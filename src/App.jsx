import { useState, useRef, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  collection, doc, addDoc, updateDoc, deleteDoc,
  onSnapshot, setDoc, getDoc
} from "firebase/firestore";
import { db } from "./firebase";

// ─── Constants ────────────────────────────────────────────────────────────────
const ROMAN = ["I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII"];
const INDO_MONTHS = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const STATUS_OPTS = ["1 - Cold","2 - Warm","3 - Hot","4 - Closed","5 - Failed","6 - Terselesaikan"];
const STATUS_COLORS = {
  "1 - Cold":{ bg:"#e6f1fb", text:"#185fa5", border:"#b5d4f4" },
  "2 - Warm":{ bg:"#faeeda", text:"#854f0b", border:"#fac775" },
  "3 - Hot":{ bg:"#fcebeb", text:"#a32d2d", border:"#f7c1c1" },
  "4 - Closed":{ bg:"#eaf3de", text:"#3b6d11", border:"#c0dd97" },
  "5 - Failed":{ bg:"#fce8e8", text:"#791f1f", border:"#f09595" },
  "6 - Terselesaikan":{ bg:"#eeedfe", text:"#3c3489", border:"#cecbf6" },
};
const DEFAULT_MASTER = {
  areas:[
    { id:1, name:"Jawa Timur", subs:["Surabaya","Malang","Sidoarjo","Gresik","Pasuruan"] },
    { id:2, name:"Bali", subs:["Denpasar","Badung","Gianyar"] },
    { id:3, name:"Kalimantan", subs:["Balikpapan","Samarinda","Banjarmasin"] },
    { id:4, name:"Sulawesi", subs:["Makassar","Manado","Kendari"] },
  ],
  pic:["Andi Santoso","Budi Pratama","Citra Dewi","Dian Rahayu","Eko Wahyudi"],
  platform:["WhatsApp","Referral","Instagram","Facebook","Website","Telepon","Email","Kunjungan Langsung","Pameran"],
  productGroup:["Outboard Engine","Inboard Engine","Generator Set","Spare Part","Service & Repair","Aksesoris"],
  purpose:["Pembelian Baru","Penggantian Unit","Servis Berkala","Survey Harga","Informasi Produk","Trade-in","Konsultasi"],
  custGroup:["1 - Dealer","2 - Fleet Owner","3 - Boat Builder","4 - Kontraktor","5 - Government","6 - Bengkel","7 - Nelayan","10 - Lainnya","12 - Makelar","13 - Rescue","14 - Private User","15 - Wisata"],
};
const LEAD_COLS = ["Lead Id","Date Started","PIC","Platform","Customer Name","Parent Name","Cust Group","Area","Sub Area","Contact","Purpose","Product Group","Value","Status","Failed Report","Date Ended","Lead Time","Notes","Year","Periode"];
const PRICE_COLS = ["Lead ID","Tanggal","PIC","Platform","Product Group","Product Name","Brand","Seller","Area - Sub Area","Price","Year","Period","Notes"];
const PAGE_SIZE = 25;

// ─── Helpers ──────────────────────────────────────────────────────────────────
function todayStr() {
  const d = new Date();
  return `${d.getDate().toString().padStart(2,"0")}/${(d.getMonth()+1).toString().padStart(2,"0")}/${d.getFullYear()}`;
}
function dateToLeadSuffix(dateStr) {
  const p = dateStr ? dateStr.split("/") : [];
  if (p.length!==3) { const n=new Date(); return `${n.getDate().toString().padStart(2,"0")}${ROMAN[n.getMonth()]}${n.getFullYear()}`; }
  return `${p[0]}${ROMAN[parseInt(p[1])-1]}${p[2]}`;
}
function dateToYear(d) { const p=d?d.split("/"):[]; return p.length===3?p[2]:new Date().getFullYear().toString(); }
function dateToPeriode(d) { const p=d?d.split("/"):[]; if(p.length!==3)return`${INDO_MONTHS[new Date().getMonth()]} ${new Date().getFullYear()}`; return`${INDO_MONTHS[parseInt(p[1])-1]} ${p[2]}`; }
function formatContact(raw) {
  const d = raw.replace(/\D/g,"");
  if (d.length<=4) return d;
  if (d.length<=8) return `${d.slice(0,4)}-${d.slice(4)}`;
  if (d.length<=12) return `${d.slice(0,4)}-${d.slice(4,8)}-${d.slice(8)}`;
  return `${d.slice(0,4)}-${d.slice(4,8)}-${d.slice(8,12)}`;
}
function normalizePhone(s) { return (s||"").replace(/\D/g,""); }
function genLeadId(leads, dateStr) { return `${(leads.length+1).toString().padStart(3,"0")}-${dateToLeadSuffix(dateStr)}`; }

// ─── Badge ────────────────────────────────────────────────────────────────────
function Badge({ text }) {
  if (!text) return <span style={{color:"#bbb",fontSize:11}}>—</span>;
  const c = STATUS_COLORS[text]||{bg:"#f1f1f1",text:"#555",border:"#ccc"};
  return <span style={{background:c.bg,color:c.text,border:`1px solid ${c.border}`,padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>{text}</span>;
}

// ─── SearchableDropdown ───────────────────────────────────────────────────────
function SearchableDropdown({ value, onChange, options=[], placeholder="Pilih...", disabled=false, error=false }) {
  const [open,setOpen]=useState(false);
  const [q,setQ]=useState("");
  const ref=useRef();
  useEffect(()=>{
    const h=e=>{if(ref.current&&!ref.current.contains(e.target)){setOpen(false);setQ("");}};
    document.addEventListener("mousedown",h);
    return()=>document.removeEventListener("mousedown",h);
  },[]);
  const filtered=options.filter(o=>(typeof o==="string"?o:o.label).toLowerCase().includes(q.toLowerCase()));
  return (
    <div ref={ref} style={{position:"relative"}}>
      <div onClick={()=>!disabled&&setOpen(o=>!o)}
        style={{padding:"8px 10px",borderRadius:6,border:`1.5px solid ${error?"#d93025":open?"#185fa5":"#d0d5dd"}`,background:disabled?"#f5f5f5":"#fff",color:value?"#1a1a1a":"#aaa",fontSize:13,cursor:disabled?"default":"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",minHeight:36,userSelect:"none"}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{value||placeholder}</span>
        <span style={{marginLeft:8,fontSize:10,color:"#888",flexShrink:0}}>{open?"▲":"▼"}</span>
      </div>
      {open&&(
        <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",border:"1.5px solid #d0d5dd",borderRadius:8,zIndex:9000,boxShadow:"0 8px 24px rgba(0,0,0,0.14)",overflow:"hidden"}}>
          <input autoFocus value={q} onChange={e=>setQ(e.target.value)} placeholder="Cari..."
            style={{padding:"8px 12px",border:"none",borderBottom:"1px solid #eee",outline:"none",fontSize:13,width:"100%",boxSizing:"border-box"}}/>
          <div style={{maxHeight:200,overflowY:"auto"}}>
            {filtered.length===0?<div style={{padding:"12px",color:"#aaa",fontSize:13,textAlign:"center"}}>Tidak ditemukan</div>
              :filtered.map((o,i)=>{
                const label=typeof o==="string"?o:o.label, val=typeof o==="string"?o:o.value;
                return(<div key={i} onClick={()=>{onChange(val);setOpen(false);setQ("");}}
                  style={{padding:"9px 12px",cursor:"pointer",fontSize:13,color:"#1a1a1a",borderBottom:"1px solid #f5f5f5"}}
                  onMouseEnter={e=>e.currentTarget.style.background="#f0f7ff"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{label}</div>);
              })
            }
          </div>
        </div>
      )}
    </div>
  );
}

// ─── CreatableDropdown ────────────────────────────────────────────────────────
function CreatableDropdown({ value, onChange, options=[], placeholder="Ketik atau pilih...", error=false }) {
  const [open,setOpen]=useState(false);
  const [q,setQ]=useState(value||"");
  const ref=useRef();
  useEffect(()=>{setQ(value||"");},[value]);
  useEffect(()=>{
    const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);};
    document.addEventListener("mousedown",h);
    return()=>document.removeEventListener("mousedown",h);
  },[]);
  const filtered=q.trim()?options.filter(o=>o.toLowerCase().includes(q.toLowerCase())):options;
  const isExact=options.some(o=>o.toLowerCase()===q.trim().toLowerCase());
  const isNew=q.trim()&&!isExact;
  return (
    <div ref={ref} style={{position:"relative"}}>
      <div style={{position:"relative"}}>
        <input value={q} onChange={e=>{setQ(e.target.value);onChange(e.target.value);setOpen(true);}}
          onFocus={()=>setOpen(true)} onBlur={()=>setTimeout(()=>setOpen(false),150)}
          placeholder={placeholder}
          style={{padding:"8px 10px",paddingRight:q.trim()?90:10,borderRadius:6,border:`1.5px solid ${error?"#d93025":open?"#185fa5":"#d0d5dd"}`,background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}}/>
        {q.trim()&&<span style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",fontSize:10,fontWeight:700,padding:"2px 7px",borderRadius:10,background:isExact?"#eaf3de":"#faeeda",color:isExact?"#2d6a11":"#854f0b",pointerEvents:"none",whiteSpace:"nowrap"}}>
          {isExact?"✓ Terdaftar":"⊕ Baru"}
        </span>}
      </div>
      {open&&(filtered.length>0||isNew)&&(
        <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",border:"1.5px solid #d0d5dd",borderRadius:8,zIndex:9000,boxShadow:"0 8px 24px rgba(0,0,0,0.14)",overflow:"hidden"}}>
          <div style={{maxHeight:180,overflowY:"auto"}}>
            {filtered.map((opt,i)=>(
              <div key={i} onMouseDown={()=>{setQ(opt);onChange(opt);setOpen(false);}}
                style={{padding:"9px 12px",cursor:"pointer",fontSize:13,color:"#1a1a1a",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",gap:8}}
                onMouseEnter={e=>e.currentTarget.style.background="#f0f7ff"}
                onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                <span style={{fontSize:12}}>🏢</span><span style={{flex:1}}>{opt}</span>
              </div>
            ))}
          </div>
          {isNew&&<div style={{padding:"9px 12px",fontSize:12,color:"#854f0b",background:"#fffbf0",borderTop:filtered.length>0?"1px solid #f5e8c0":"none"}}>
            ⊕ Tambah baru: <strong>"{q.trim()}"</strong>
          </div>}
        </div>
      )}
    </div>
  );
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function Modal({ title, onClose, children, wide=false }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}
      onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"#fff",borderRadius:12,width:wide?"min(780px,96vw)":"min(620px,96vw)",maxHeight:"92vh",boxShadow:"0 20px 60px rgba(0,0,0,0.22)",display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"15px 20px",borderBottom:"1px solid #e8e8e8",background:"#f9fafb",flexShrink:0}}>
          <span style={{fontWeight:700,fontSize:15,color:"#1a1a1a"}}>{title}</span>
          <button onClick={onClose} style={{background:"#eee",border:"none",cursor:"pointer",fontSize:16,color:"#555",width:28,height:28,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:600}}>×</button>
        </div>
        <div style={{padding:20,overflowY:"auto",flex:1}}>{children}</div>
      </div>
    </div>
  );
}

// ─── LeadForm ─────────────────────────────────────────────────────────────────
function LeadForm({ initial={}, leads=[], master, onSave, onCancel, isEdit=false }) {
  const today=todayStr();
  const [form,setForm]=useState(()=>{
    if (!isEdit) {
      const lid=genLeadId(leads,today);
      return {"Lead Id":lid,"Date Started":today,"Year":dateToYear(today),"Periode":dateToPeriode(today),"Date Ended":"","PIC":"","Platform":"","Customer Name":"","Parent Name":"","Cust Group":"","Area":"","Sub Area":"","Contact":"","Purpose":"","Product Group":"","Value":"","Status":"1 - Cold","Failed Report":"","Lead Time":"","Notes":""};
    }
    return {...initial};
  });
  const [errors,setErrors]=useState({});
  const [contactMatch,setContactMatch]=useState(null);

  const set=(k,v)=>setForm(p=>({...p,[k]:v}));

  const parentNameOptions=useMemo(()=>{
    const names=leads.map(l=>(l["Parent Name"]||"").trim()).filter(Boolean);
    return [...new Set(names)].sort();
  },[leads]);

  const handleContactChange=(raw)=>{
    const digits=raw.replace(/\D/g,"");
    set("Contact",formatContact(digits));
    if (digits.length>=10) {
      const match=leads.find(l=>{ if(isEdit&&l["Lead Id"]===initial["Lead Id"])return false; return normalizePhone(l["Contact"])===digits; });
      setContactMatch(match?{name:match["Customer Name"]||"",parentName:match["Parent Name"]||""}:null);
    } else setContactMatch(null);
  };

  const areaOptions=useMemo(()=>{
    const opts=[];
    master.areas.forEach(a=>{
      if(a.subs&&a.subs.length) a.subs.forEach(s=>opts.push({label:`${a.name} › ${s}`,value:`${a.name}||${s}`}));
      else opts.push({label:a.name,value:`${a.name}||`});
    });
    return opts;
  },[master.areas]);
  const handleAreaSelect=(val)=>{const[area,sub]=val.split("||");setForm(p=>({...p,Area:area,"Sub Area":sub||""}));};
  const areaDisplay=form["Area"]?(form["Sub Area"]?`${form["Area"]} › ${form["Sub Area"]}`:form["Area"]):"";

  const validate=()=>{
    const e={};
    if(!form["Customer Name"].trim()) e["Customer Name"]="Wajib diisi";
    if(!form["Contact"].trim()) e["Contact"]="Wajib diisi";
    else { const d=normalizePhone(form["Contact"]); if(d.length<10||d.length>13) e["Contact"]="Format: 0XXX-XXXX-XXXX"; }
    if(!form["Area"]) e["Area"]="Wajib dipilih";
    if(!form["Status"]) e["Status"]="Wajib dipilih";
    setErrors(e); return Object.keys(e).length===0;
  };

  const handleSave=()=>{
    if(!validate()) return;
    const saved={...form};
    if(isEdit) saved["Date Ended"]=today;
    onSave(saved);
  };

  const field=(label,content,required=false,errKey=null)=>(
    <div style={{display:"flex",flexDirection:"column",gap:4}}>
      <label style={{fontSize:11,color:"#555",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.04em"}}>
        {label}{required&&<span style={{color:"#d93025",marginLeft:2}}>*</span>}
      </label>
      {content}
      {errors[errKey||label]&&<span style={{fontSize:11,color:"#d93025"}}>{errors[errKey||label]}</span>}
    </div>
  );
  const inp=(k,opts={})=>(
    <input value={form[k]||""} onChange={e=>set(k,e.target.value)} readOnly={opts.readOnly} placeholder={opts.placeholder||""}
      style={{padding:"8px 10px",borderRadius:6,border:`1.5px solid ${errors[k]?"#d93025":"#d0d5dd"}`,background:opts.readOnly?"#f5f5f5":"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}}
      onFocus={e=>{if(!opts.readOnly)e.target.style.borderColor="#185fa5";}}
      onBlur={e=>{if(!opts.readOnly)e.target.style.borderColor=errors[k]?"#d93025":"#d0d5dd";}}/>
  );
  const dd=(k,options,placeholder)=>(
    <SearchableDropdown value={form[k]||""} onChange={v=>set(k,v)} options={options} placeholder={placeholder||"Pilih..."} error={!!errors[k]}/>
  );

  return (
    <div>
      <div style={{background:"#f0f7ff",borderRadius:8,padding:"10px 14px",marginBottom:16,display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10}}>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Lead ID</span><span style={{fontFamily:"monospace",fontWeight:700,color:"#185fa5",fontSize:13}}>{form["Lead Id"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Tanggal Mulai</span><span style={{fontSize:13,color:"#333"}}>{form["Date Started"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Periode</span><span style={{fontSize:13,color:"#333"}}>{form["Periode"]}</span></div>
        <div>
          <span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Tgl. Terakhir Edit</span>
          {isEdit?<span style={{fontSize:12,fontWeight:700,color:"#2d6a11"}}>✓ {today}</span>:<span style={{fontSize:12,color:"#aaa",fontStyle:"italic"}}>Auto saat di-edit</span>}
        </div>
      </div>

      {contactMatch&&(
        <div style={{background:"#fffbe6",border:"1.5px solid #f5c100",borderRadius:8,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:18}}>👤</span>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:700,color:"#7a4f00"}}>Nomor ini sudah terdaftar!</div>
            <div style={{fontSize:12,color:"#555",marginTop:2}}><strong>{contactMatch.name}</strong>{contactMatch.parentName?` · ${contactMatch.parentName}`:""}</div>
          </div>
          <button onClick={()=>{setForm(p=>({...p,"Customer Name":contactMatch.name,"Parent Name":contactMatch.parentName}));setContactMatch(null);}} style={{padding:"6px 14px",borderRadius:6,border:"none",background:"#f5c100",color:"#7a4f00",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>✓ Gunakan Data Ini</button>
          <button onClick={()=>setContactMatch(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#aaa",fontSize:18,lineHeight:1,padding:"0 4px"}}>×</button>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px 16px",maxHeight:"50vh",overflowY:"auto",paddingRight:4,paddingBottom:4}}>
        {field("Contact",(
          <input value={form["Contact"]||""} onChange={e=>handleContactChange(e.target.value)} placeholder="08XX-XXXX-XXXX" maxLength={14}
            style={{padding:"8px 10px",borderRadius:6,border:`1.5px solid ${errors["Contact"]?"#d93025":"#d0d5dd"}`,background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"}}
            onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor=errors["Contact"]?"#d93025":"#d0d5dd"}/>
        ),true,"Contact")}
        {field("Customer Name",inp("Customer Name",{placeholder:"Nama pelanggan"}),true)}
        {field("Parent Name",(
          <CreatableDropdown value={form["Parent Name"]||""} onChange={v=>set("Parent Name",v)} options={parentNameOptions} placeholder="Nama perusahaan / merk..."/>
        ))}
        {field("Cust Group",dd("Cust Group",master.custGroup,"Pilih Grup..."))}
        {field("Area (Area › Sub Area)",(
          <SearchableDropdown value={areaDisplay} onChange={handleAreaSelect} options={areaOptions} placeholder="Pilih Area › Sub Area..." error={!!errors["Area"]}/>
        ),true,"Area")}
        {field("PIC",dd("PIC",master.pic,"Pilih PIC..."))}
        {field("Platform",dd("Platform",master.platform,"Pilih Platform..."))}
        {field("Purpose",dd("Purpose",master.purpose,"Pilih Purpose..."))}
        {field("Product Group",dd("Product Group",master.productGroup,"Pilih Produk..."))}
        {field("Value",(
          <input value={form["Value"]||""} onChange={e=>set("Value",e.target.value.replace(/\D/g,""))} placeholder="Nilai transaksi (Rp)"
            style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"}}
            onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>
        ))}
        {field("Status",dd("Status",STATUS_OPTS,"Pilih Status..."),true)}
        {form["Status"]==="5 - Failed"&&<div style={{gridColumn:"1/-1"}}>{field("Failed Report",inp("Failed Report",{placeholder:"Alasan gagal"}))}</div>}
        <div style={{gridColumn:"1/-1"}}>{field("Notes",
          <textarea value={form["Notes"]||""} onChange={e=>set("Notes",e.target.value)} placeholder="Catatan tambahan..."
            style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",resize:"vertical",minHeight:56,fontFamily:"inherit"}}
            onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>
        )}</div>
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:14,paddingTop:14,borderTop:"1px solid #e8e8e8",flexShrink:0}}>
        <button onClick={onCancel} style={{padding:"9px 20px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#f5f5f5",color:"#444",cursor:"pointer",fontSize:13,fontWeight:500}}>Batal</button>
        <button onClick={handleSave} style={{padding:"9px 20px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>
          {isEdit?"💾 Simpan Perubahan":"➕ Tambah Lead"}
        </button>
      </div>
    </div>
  );
}

// ─── PriceForm ────────────────────────────────────────────────────────────────
function PriceForm({ initial={}, master, onSave, onCancel, isEdit=false }) {
  const today=todayStr();
  const [form,setForm]=useState(()=>isEdit?{...initial}:{
    "Lead ID":"","Tanggal":today,"PIC":"","Platform":"","Product Group":"","Product Name":"","Brand":"","Seller":"","Area - Sub Area":"","Price":"","Year":dateToYear(today),"Period":dateToPeriode(today),"Notes":""
  });
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const areaOptions=useMemo(()=>{
    const opts=[];
    master.areas.forEach(a=>{
      if(a.subs&&a.subs.length) a.subs.forEach(s=>opts.push({label:`${a.name} › ${s}`,value:`${a.name} › ${s}`}));
      else opts.push({label:a.name,value:a.name});
    });
    return opts;
  },[master.areas]);
  const inp=(k,opts={})=>(
    <input value={form[k]||""} onChange={e=>set(k,e.target.value)} readOnly={opts.readOnly} placeholder={opts.placeholder||""}
      style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:opts.readOnly?"#f5f5f5":"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}}
      onFocus={e=>{if(!opts.readOnly)e.target.style.borderColor="#185fa5";}}
      onBlur={e=>{if(!opts.readOnly)e.target.style.borderColor="#d0d5dd";}}/>
  );
  const dd=(k,options,placeholder)=>(<SearchableDropdown value={form[k]||""} onChange={v=>set(k,v)} options={options} placeholder={placeholder||"Pilih..."}/>);
  const field=(label,content)=>(<div style={{display:"flex",flexDirection:"column",gap:4}}><label style={{fontSize:11,color:"#555",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.04em"}}>{label}</label>{content}</div>);
  return (
    <div>
      <div style={{background:"#f0f7ff",borderRadius:8,padding:"10px 14px",marginBottom:14,display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase"}}>Tanggal</span><span style={{fontSize:13,color:"#333"}}>{form["Tanggal"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase"}}>Year</span><span style={{fontSize:13,color:"#333"}}>{form["Year"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase"}}>Period</span><span style={{fontSize:13,color:"#333"}}>{form["Period"]}</span></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px 16px",maxHeight:"50vh",overflowY:"auto",paddingRight:4,paddingBottom:4}}>
        {field("Lead ID",inp("Lead ID",{placeholder:"Referensi Lead ID"}))}
        {field("PIC",dd("PIC",master.pic,"Pilih PIC..."))}
        {field("Platform",dd("Platform",master.platform,"Pilih Platform..."))}
        {field("Product Group",dd("Product Group",master.productGroup,"Pilih Product Group..."))}
        {field("Product Name",inp("Product Name",{placeholder:"Nama produk"}))}
        {field("Brand",inp("Brand",{placeholder:"Merek"}))}
        {field("Seller",inp("Seller",{placeholder:"Nama seller / dealer"}))}
        {field("Area - Sub Area",dd("Area - Sub Area",areaOptions,"Pilih Area..."))}
        {field("Price",(
          <input value={form["Price"]||""} onChange={e=>set("Price",e.target.value.replace(/\D/g,""))} placeholder="Harga (Rp)"
            style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"}}
            onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>
        ))}
        <div style={{gridColumn:"1/-1"}}>{field("Notes",
          <textarea value={form["Notes"]||""} onChange={e=>set("Notes",e.target.value)} placeholder="Catatan..."
            style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",resize:"vertical",minHeight:56,fontFamily:"inherit"}}
            onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>
        )}</div>
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:14,paddingTop:14,borderTop:"1px solid #e8e8e8"}}>
        <button onClick={onCancel} style={{padding:"9px 20px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#f5f5f5",color:"#444",cursor:"pointer",fontSize:13,fontWeight:500}}>Batal</button>
        <button onClick={()=>onSave(form)} style={{padding:"9px 20px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>{isEdit?"💾 Simpan":"➕ Tambah"}</button>
      </div>
    </div>
  );
}

// ─── AdminPanel ────────────────────────────────────────────────────────────────
function AdminPanel({ master, setMasterField, onBack }) {
  const [sec,setSec]=useState("areas");
  const [newAreaName,setNewAreaName]=useState("");
  const [newSub,setNewSub]=useState({});
  const [newItem,setNewItem]=useState("");
  const [importMsg,setImportMsg]=useState(null);
  const areaImportRef=useRef();

  const SECS=[{id:"areas",label:"📍 Area & Sub Area"},{id:"pic",label:"👤 PIC / Sales"},{id:"platform",label:"📱 Platform"},{id:"productGroup",label:"⚙️ Product Group"},{id:"purpose",label:"🎯 Purpose"},{id:"custGroup",label:"🏢 Customer Group"}];

  const addArea=()=>{ if(!newAreaName.trim())return; setMasterField("areas",[...master.areas,{id:Date.now(),name:newAreaName.trim(),subs:[]}]); setNewAreaName(""); };
  const removeArea=id=>setMasterField("areas",master.areas.filter(a=>a.id!==id));
  const addSub=aId=>{ const s=(newSub[aId]||"").trim(); if(!s)return; setMasterField("areas",master.areas.map(a=>a.id===aId?{...a,subs:[...a.subs,s]}:a)); setNewSub(p=>({...p,[aId]:""})); };
  const removeSub=(aId,s)=>setMasterField("areas",master.areas.map(a=>a.id===aId?{...a,subs:a.subs.filter(x=>x!==s)}:a));
  const addToList=k=>{ if(!newItem.trim()||master[k].includes(newItem.trim()))return; setMasterField(k,[...master[k],newItem.trim()]); setNewItem(""); };
  const removeFromList=(k,item)=>setMasterField(k,master[k].filter(i=>i!==item));

  const downloadAreaTemplate=()=>{
    const data=[["Area","Sub Area"],["Jawa Timur","Surabaya"],["Jawa Timur","Malang"],["Bali","Denpasar"]];
    const ws=XLSX.utils.aoa_to_sheet(data);
    ws["!cols"]=[{wch:20},{wch:20}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Area Template");
    XLSX.writeFile(wb,"Template_Area_SubArea.xlsx");
  };

  const handleAreaImport=e=>{
    const file=e.target.files[0]; if(!file)return;
    const reader=new FileReader();
    reader.onload=evt=>{
      try {
        const wb=XLSX.read(evt.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:false}).filter(r=>r.some(c=>c));
        const start=rows.length>0&&String(rows[0][0]||"").toLowerCase()==="area"?1:0;
        const areaMap={};
        rows.slice(start).forEach(row=>{
          const area=String(row[0]||"").trim(), sub=String(row[1]||"").trim();
          if(!area)return;
          if(!areaMap[area])areaMap[area]=new Set();
          if(sub)areaMap[area].add(sub);
        });
        const existing=[...master.areas];
        Object.entries(areaMap).forEach(([areaName,subsSet])=>{
          const idx=existing.findIndex(a=>a.name.toLowerCase()===areaName.toLowerCase());
          if(idx>=0){existing[idx]={...existing[idx],subs:[...new Set([...existing[idx].subs,...subsSet])]};}
          else{existing.push({id:Date.now()+Math.random(),name:areaName,subs:[...subsSet]});}
        });
        setMasterField("areas",existing);
        const ta=Object.keys(areaMap).length, ts=Object.values(areaMap).reduce((s,v)=>s+v.size,0);
        setImportMsg({ok:true,text:`Berhasil import ${ta} area dan ${ts} sub area`});
        setTimeout(()=>setImportMsg(null),4000);
      } catch(err){ setImportMsg({ok:false,text:"Gagal baca file: "+err.message}); setTimeout(()=>setImportMsg(null),4000); }
    };
    reader.readAsBinaryString(file); e.target.value="";
  };

  return (
    <div style={{minHeight:"100vh",background:"#f4f6f9",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif"}}>
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"14px 24px",display:"flex",alignItems:"center",gap:12}}>
        <button onClick={onBack} style={{background:"#f0f7ff",border:"1px solid #b5d4f4",borderRadius:7,padding:"7px 14px",cursor:"pointer",color:"#185fa5",fontSize:13,fontWeight:600}}>← Dashboard</button>
        <div><div style={{fontWeight:700,fontSize:16,color:"#1a1a1a"}}>⚙️ Admin — Master Data</div><div style={{fontSize:12,color:"#888"}}>Perubahan tersimpan otomatis ke Firebase</div></div>
        <span style={{marginLeft:"auto",fontSize:12,background:"#eaf3de",color:"#2d6a11",padding:"4px 10px",borderRadius:20,fontWeight:600}}>🔥 Firebase Connected</span>
      </div>
      <div style={{display:"flex",gap:0,padding:"20px 24px"}}>
        <div style={{width:196,flexShrink:0,marginRight:18}}>
          {SECS.map(s=>(
            <button key={s.id} onClick={()=>setSec(s.id)} style={{display:"block",width:"100%",textAlign:"left",padding:"10px 14px",borderRadius:8,border:"none",background:sec===s.id?"#185fa5":"transparent",color:sec===s.id?"#fff":"#444",fontSize:13,fontWeight:sec===s.id?600:400,cursor:"pointer",marginBottom:4}}>
              {s.label}
            </button>
          ))}
        </div>
        <div style={{flex:1,background:"#fff",borderRadius:12,border:"1px solid #e5e7eb",padding:20}}>
          {sec==="areas"?(
            <div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{fontWeight:700,fontSize:15}}>📍 Area & Sub Area</div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={downloadAreaTemplate} style={{padding:"7px 13px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#fff",color:"#555",cursor:"pointer",fontSize:12,fontWeight:500}}>📄 Unduh Template</button>
                  <input ref={areaImportRef} type="file" accept=".xlsx,.xls" onChange={handleAreaImport} style={{display:"none"}}/>
                  <button onClick={()=>areaImportRef.current?.click()} style={{padding:"7px 13px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600}}>⬆️ Import Excel</button>
                </div>
              </div>
              <div style={{background:"#fffbf0",border:"1px solid #f5e8c0",borderRadius:8,padding:"9px 13px",marginBottom:14,fontSize:12,color:"#7a4f00"}}>
                💡 <strong>Format:</strong> Kolom A = Area, Kolom B = Sub Area. Satu baris per kombinasi.
              </div>
              {importMsg&&<div style={{background:importMsg.ok?"#eaf3de":"#fce8e8",border:`1px solid ${importMsg.ok?"#b3d97a":"#f0a0a0"}`,borderRadius:8,padding:"9px 13px",marginBottom:14,fontSize:13,fontWeight:500,color:importMsg.ok?"#2d6a11":"#a32d2d"}}>{importMsg.ok?"✓":"✕"} {importMsg.text}</div>}
              <div style={{display:"flex",gap:8,marginBottom:20}}>
                <input value={newAreaName} onChange={e=>setNewAreaName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addArea()} placeholder="Nama Area baru..."
                  style={{flex:1,padding:"8px 12px",borderRadius:7,border:"1.5px solid #d0d5dd",fontSize:13,outline:"none"}}/>
                <button onClick={addArea} style={{padding:"8px 16px",borderRadius:7,border:"none",background:"#eaf3de",color:"#2d6a11",cursor:"pointer",fontSize:13,fontWeight:600}}>+ Tambah Area</button>
              </div>
              {master.areas.map(area=>(
                <div key={area.id} style={{border:"1px solid #e5e7eb",borderRadius:10,marginBottom:12,overflow:"hidden"}}>
                  <div style={{background:"#f8fafc",padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:600,fontSize:14}}>🗺️ {area.name}</span>
                    <button onClick={()=>removeArea(area.id)} style={{background:"#fce8e8",border:"none",color:"#d93025",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:600}}>Hapus</button>
                  </div>
                  <div style={{padding:"12px 14px"}}>
                    <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>
                      {area.subs.map(s=>(
                        <span key={s} style={{background:"#e6f1fb",color:"#185fa5",borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:500,display:"flex",alignItems:"center",gap:4}}>
                          {s}<button onClick={()=>removeSub(area.id,s)} style={{background:"none",border:"none",cursor:"pointer",color:"#a32d2d",fontSize:14,lineHeight:1,padding:0}}>×</button>
                        </span>
                      ))}
                      {area.subs.length===0&&<span style={{color:"#bbb",fontSize:12}}>Belum ada sub area</span>}
                    </div>
                    <div style={{display:"flex",gap:8}}>
                      <input value={newSub[area.id]||""} onChange={e=>setNewSub(p=>({...p,[area.id]:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addSub(area.id)} placeholder="Sub Area baru..."
                        style={{flex:1,padding:"7px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",fontSize:13,outline:"none"}}/>
                      <button onClick={()=>addSub(area.id)} style={{padding:"7px 14px",borderRadius:6,border:"none",background:"#eaf3de",color:"#3b6d11",cursor:"pointer",fontSize:13,fontWeight:600}}>+ Sub Area</button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          ):(
            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>{SECS.find(s=>s.id===sec)?.label}</div>
              <div style={{display:"flex",gap:8,marginBottom:16}}>
                <input value={newItem} onChange={e=>setNewItem(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addToList(sec)} placeholder="Tambah item baru..."
                  style={{flex:1,padding:"8px 12px",borderRadius:7,border:"1.5px solid #d0d5dd",fontSize:13,outline:"none"}}/>
                <button onClick={()=>addToList(sec)} style={{padding:"8px 16px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>+ Tambah</button>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {master[sec].map((item,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 14px",border:"1px solid #e5e7eb",borderRadius:8,background:"#fafbfc"}}>
                    <span style={{fontSize:13}}>{item}</span>
                    <button onClick={()=>removeFromList(sec,item)} style={{background:"#fce8e8",border:"none",color:"#d93025",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:600}}>Hapus</button>
                  </div>
                ))}
                {master[sec].length===0&&<div style={{color:"#bbb",fontSize:13,padding:"20px 0",textAlign:"center"}}>Belum ada data.</div>}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── LoginScreen ───────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0f3d6e 0%,#185fa5 60%,#1e7fc4 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px 36px",width:"min(400px,100%)",boxShadow:"0 24px 80px rgba(0,0,0,0.22)",textAlign:"center"}}>
        <div style={{fontSize:48,marginBottom:8}}>⚓</div>
        <div style={{fontWeight:800,fontSize:22,color:"#1a1a1a",marginBottom:4}}>Lead Conversion Dashboard</div>
        <div style={{fontSize:13,color:"#888",marginBottom:4}}>Marine Engine Sales Tracker</div>
        <div style={{fontSize:12,color:"#2d6a11",background:"#eaf3de",padding:"4px 12px",borderRadius:20,display:"inline-block",marginBottom:28,fontWeight:600}}>🔥 Powered by Firebase</div>
        <div style={{fontWeight:600,fontSize:14,color:"#555",marginBottom:16,textAlign:"left"}}>Masuk sebagai:</div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={()=>onLogin("admin")} style={{padding:"14px 20px",borderRadius:10,border:"none",background:"linear-gradient(135deg,#185fa5,#1e7fc4)",color:"#fff",cursor:"pointer",fontSize:15,fontWeight:700,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:24}}>🔐</span>
            <div style={{textAlign:"left"}}><div>Admin</div><div style={{fontSize:11,fontWeight:400,opacity:0.85}}>Kelola master data & semua lead</div></div>
          </button>
          <button onClick={()=>onLogin("sales")} style={{padding:"14px 20px",borderRadius:10,border:"2px solid #185fa5",background:"#f0f7ff",color:"#185fa5",cursor:"pointer",fontSize:15,fontWeight:700,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:24}}>💼</span>
            <div style={{textAlign:"left"}}><div>Sales</div><div style={{fontSize:11,fontWeight:400,opacity:0.8}}>Input & kelola lead pelanggan</div></div>
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── GenericTable ──────────────────────────────────────────────────────────────
function GenericTable({ rows, visibleCols, page, totalPages, filtered, onEdit, onDelete, role, onPageChange }) {
  return (
    <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,overflow:"hidden"}}>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5}}>
          <thead>
            <tr style={{background:"#f8fafc",borderBottom:"1px solid #e5e7eb"}}>
              <th style={{padding:"10px 10px",textAlign:"center",color:"#aaa",fontWeight:600,fontSize:10,width:36}}>#</th>
              {visibleCols.map(col=>(
                <th key={col} style={{padding:"10px 12px",textAlign:"left",color:"#666",fontWeight:700,fontSize:10,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.05em"}}>{col}</th>
              ))}
              <th style={{padding:"10px 12px",color:"#aaa",fontSize:10,width:80,textAlign:"center",textTransform:"uppercase",letterSpacing:"0.05em"}}>Aksi</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row,i)=>{
              const globalIdx=(page-1)*PAGE_SIZE+i;
              return (
                <tr key={row._id||i} style={{borderBottom:"1px solid #f0f0f0"}}
                  onMouseEnter={e=>e.currentTarget.style.background="#fafbfc"}
                  onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={{padding:"9px 10px",textAlign:"center",color:"#bbb",fontSize:11}}>{globalIdx+1}</td>
                  {visibleCols.map(col=>(
                    <td key={col} style={{padding:"9px 12px",color:"#1a1a1a",maxWidth:160}}>
                      {col==="Status"?<Badge text={row[col]}/>
                       :col==="Lead Id"||col==="Lead ID"?<span style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#185fa5"}}>{row[col]}</span>
                       :col==="Value"||col==="Price"?<span style={{fontWeight:700,color:row[col]?"#2d6a11":"#bbb"}}>{row[col]?`Rp ${Number(row[col]).toLocaleString("id-ID")}`:"—"}</span>
                       :col==="Cust Group"?<span style={{fontSize:11,fontWeight:500,color:"#534ab7"}}>{row[col]||"—"}</span>
                       :<span style={{display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:150}} title={String(row[col]||"")}>{row[col]||<span style={{color:"#ccc"}}>—</span>}</span>
                      }
                    </td>
                  ))}
                  <td style={{padding:"9px 10px",textAlign:"center",whiteSpace:"nowrap"}}>
                    <button onClick={()=>onEdit(row)} style={{background:"#e8f0fb",border:"none",cursor:"pointer",color:"#185fa5",padding:"5px 8px",borderRadius:6,fontSize:13,marginRight:4}}>✏️</button>
                    {role==="admin"&&<button onClick={()=>onDelete(row._id)} style={{background:"#fce8e8",border:"none",cursor:"pointer",color:"#a32d2d",padding:"5px 8px",borderRadius:6,fontSize:13}}>🗑️</button>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {totalPages>1&&(
        <div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:6,padding:"12px 16px",borderTop:"1px solid #f0f0f0"}}>
          <button onClick={()=>onPageChange(p=>Math.max(1,p-1))} disabled={page===1} style={{padding:"6px 12px",borderRadius:6,border:"1px solid #e0e0e0",background:"#fff",cursor:"pointer",fontSize:12,color:"#555"}}>‹</button>
          {Array.from({length:Math.min(7,totalPages)},(_,i)=>{const s=Math.max(1,Math.min(page-3,totalPages-6));return s+i;}).filter(p=>p>=1&&p<=totalPages).map(p2=>(
            <button key={p2} onClick={()=>onPageChange(p2)} style={{padding:"6px 10px",minWidth:32,borderRadius:6,border:page===p2?"none":"1px solid #e0e0e0",background:page===p2?"#185fa5":"#fff",color:page===p2?"#fff":"#555",cursor:"pointer",fontSize:12,fontWeight:page===p2?700:400}}>{p2}</button>
          ))}
          <button onClick={()=>onPageChange(p=>Math.min(totalPages,p+1))} disabled={page===totalPages} style={{padding:"6px 12px",borderRadius:6,border:"1px solid #e0e0e0",background:"#fff",cursor:"pointer",fontSize:12,color:"#555"}}>›</button>
          <span style={{fontSize:12,color:"#aaa",marginLeft:4}}>{(page-1)*PAGE_SIZE+1}–{Math.min(page*PAGE_SIZE,filtered.length)} / {filtered.length}</span>
        </div>
      )}
    </div>
  );
}

// ─── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [role,setRole]=useState(null);
  const [showAdmin,setShowAdmin]=useState(false);
  const [master,setMaster]=useState(DEFAULT_MASTER);
  const [masterLoading,setMasterLoading]=useState(true);
  const [activeTab,setActiveTab]=useState("lead");

  const [leads,setLeads]=useState([]);
  const [leadsLoading,setLeadsLoading]=useState(true);
  const [leadSearch,setLeadSearch]=useState("");
  const [filterStatus,setFilterStatus]=useState("all");
  const [filterGroup,setFilterGroup]=useState("all");
  const [leadPage,setLeadPage]=useState(1);
  const [leadModal,setLeadModal]=useState(null);
  const [leadDeleteId,setLeadDeleteId]=useState(null);

  const [prices,setPrices]=useState([]);
  const [pricesLoading,setPricesLoading]=useState(true);
  const [priceSearch,setPriceSearch]=useState("");
  const [pricePage,setPricePage]=useState(1);
  const [priceModal,setPriceModal]=useState(null);
  const [priceDeleteId,setPriceDeleteId]=useState(null);

  const [toast,setToast]=useState(null);
  const [dbError,setDbError]=useState(null);
  const importRef=useRef();

  const showToast=(msg,ok=true)=>{setToast({msg,ok});setTimeout(()=>setToast(null),3200);};

  // ── Firebase: Load master data ──
  useEffect(()=>{
    const ref=doc(db,"config","masterData");
    getDoc(ref).then(snap=>{
      if(snap.exists()){
        const data=snap.data();
        setMaster(prev=>({...DEFAULT_MASTER,...prev,...data}));
      } else {
        // First time: seed default master data
        setDoc(ref,DEFAULT_MASTER).catch(()=>{});
      }
      setMasterLoading(false);
    }).catch(err=>{
      console.error("Firebase master error:",err);
      setDbError("Tidak dapat terhubung ke Firebase. Periksa konfigurasi di firebase.js");
      setMasterLoading(false);
    });
  },[]);

  // ── Firebase: Real-time leads listener ──
  useEffect(()=>{
    const unsub=onSnapshot(collection(db,"leads"),(snap)=>{
      const data=snap.docs.map(d=>({_id:d.id,...d.data()}));
      // Sort by Date Started desc
      data.sort((a,b)=>(b["Date Started"]||"").localeCompare(a["Date Started"]||""));
      setLeads(data);
      setLeadsLoading(false);
    },(err)=>{
      console.error("Leads listener error:",err);
      setLeadsLoading(false);
    });
    return()=>unsub();
  },[]);

  // ── Firebase: Real-time prices listener ──
  useEffect(()=>{
    const unsub=onSnapshot(collection(db,"prices"),(snap)=>{
      const data=snap.docs.map(d=>({_id:d.id,...d.data()}));
      data.sort((a,b)=>(b["Tanggal"]||"").localeCompare(a["Tanggal"]||""));
      setPrices(data);
      setPricesLoading(false);
    },(err)=>{
      console.error("Prices listener error:",err);
      setPricesLoading(false);
    });
    return()=>unsub();
  },[]);

  // ── Save master field to Firebase ──
  const setMasterField=useCallback((field,value)=>{
    const updated={...master,[field]:value};
    setMaster(updated);
    setDoc(doc(db,"config","masterData"),updated).catch(err=>{
      showToast("Gagal simpan ke Firebase: "+err.message,false);
    });
  },[master]);

  // ── Import ──
  const handleImport=useCallback(e=>{
    const file=e.target.files[0]; if(!file)return;
    const reader=new FileReader();
    reader.onload=async evt=>{
      try {
        const wb=XLSX.read(evt.target.result,{type:"binary",cellDates:true});
        let imported=0;
        if(wb.Sheets["Lead Tracker"]){
          const raw=XLSX.utils.sheet_to_json(wb.Sheets["Lead Tracker"],{header:1,raw:false});
          const hdr=raw[1]||[];
          const rows=raw.slice(2).filter(r=>r.some(c=>c!=null&&String(c).trim()!=="")).map(row=>{
            const obj={}; LEAD_COLS.forEach(col=>{const idx=hdr.findIndex(h=>String(h||"").trim()===col);obj[col]=idx>=0?(row[idx]??""):"";});return obj;
          });
          for(const row of rows){ await addDoc(collection(db,"leads"),row); imported++; }
        }
        if(wb.Sheets["Product Price"]){
          const raw=XLSX.utils.sheet_to_json(wb.Sheets["Product Price"],{header:1,raw:false});
          const hdr=raw[1]||[];
          const rows=raw.slice(2).filter(r=>r.some(c=>c!=null&&String(c).trim()!=="")).map(row=>{
            const obj={}; PRICE_COLS.forEach(col=>{const idx=hdr.findIndex(h=>String(h||"").trim()===col);obj[col]=idx>=0?(row[idx]??""):"";});return obj;
          });
          for(const row of rows){ await addDoc(collection(db,"prices"),row); imported++; }
        }
        showToast(`${imported} baris berhasil diimpor ke Firebase`);
      } catch(err){showToast("Gagal import: "+err.message,false);}
    };
    reader.readAsBinaryString(file); e.target.value="";
  },[]);

  // ── Export ──
  const handleExport=useCallback(()=>{
    if(!leads.length&&!prices.length){showToast("Tidak ada data!",false);return;}
    const clean=arr=>arr.map(({_id,...rest})=>rest);
    const wb=XLSX.utils.book_new();
    if(leads.length){const ws=XLSX.utils.json_to_sheet(clean(leads),{header:LEAD_COLS});XLSX.utils.book_append_sheet(wb,ws,"Lead Tracker");}
    if(prices.length){const ws=XLSX.utils.json_to_sheet(clean(prices),{header:PRICE_COLS});XLSX.utils.book_append_sheet(wb,ws,"Product Price");}
    XLSX.writeFile(wb,"LeadConversion_export.xlsx");
    showToast("File berhasil didownload!");
  },[leads,prices]);

  // ── Lead CRUD ──
  const handleSaveLead=async(form,isEdit,docId)=>{
    const {_id,...data}=form;
    try {
      if(isEdit){ await updateDoc(doc(db,"leads",docId),data); showToast("Lead diperbarui!"); }
      else { await addDoc(collection(db,"leads"),data); showToast("Lead baru ditambahkan!"); }
      setLeadModal(null);
    } catch(err){showToast("Gagal simpan: "+err.message,false);}
  };
  const handleDeleteLead=async id=>{
    try{ await deleteDoc(doc(db,"leads",id)); showToast("Lead dihapus."); setLeadDeleteId(null); }
    catch(err){showToast("Gagal hapus: "+err.message,false);}
  };

  // ── Price CRUD ──
  const handleSavePrice=async(form,isEdit,docId)=>{
    const {_id,...data}=form;
    try {
      if(isEdit){ await updateDoc(doc(db,"prices",docId),data); showToast("Data harga diperbarui!"); }
      else { await addDoc(collection(db,"prices"),data); showToast("Data harga ditambahkan!"); }
      setPriceModal(null);
    } catch(err){showToast("Gagal simpan: "+err.message,false);}
  };
  const handleDeletePrice=async id=>{
    try{ await deleteDoc(doc(db,"prices",id)); showToast("Data harga dihapus."); setPriceDeleteId(null); }
    catch(err){showToast("Gagal hapus: "+err.message,false);}
  };

  const statusOptions=useMemo(()=>[...new Set(leads.map(r=>r["Status"]).filter(Boolean))].sort(),[leads]);
  const groupOptions=useMemo(()=>[...new Set(leads.map(r=>r["Cust Group"]).filter(Boolean))].sort(),[leads]);

  const filteredLeads=useMemo(()=>{
    let d=leads;
    if(leadSearch.trim()){const q=leadSearch.toLowerCase();d=d.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q)));}
    if(filterStatus!=="all") d=d.filter(r=>r["Status"]===filterStatus);
    if(filterGroup!=="all") d=d.filter(r=>r["Cust Group"]===filterGroup);
    return d;
  },[leads,leadSearch,filterStatus,filterGroup]);

  const filteredPrices=useMemo(()=>{
    if(!priceSearch.trim()) return prices;
    const q=priceSearch.toLowerCase();
    return prices.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q)));
  },[prices,priceSearch]);

  const leadTotalPages=Math.max(1,Math.ceil(filteredLeads.length/PAGE_SIZE));
  const priceTotalPages=Math.max(1,Math.ceil(filteredPrices.length/PAGE_SIZE));
  const paginatedLeads=filteredLeads.slice((leadPage-1)*PAGE_SIZE,leadPage*PAGE_SIZE);
  const paginatedPrices=filteredPrices.slice((pricePage-1)*PAGE_SIZE,pricePage*PAGE_SIZE);

  const stats=useMemo(()=>{
    const t=leads.length,cl=leads.filter(r=>r["Status"]==="4 - Closed").length,
      ht=leads.filter(r=>r["Status"]==="3 - Hot").length,fl=leads.filter(r=>r["Status"]==="5 - Failed").length,
      tv=leads.reduce((s,r)=>s+(parseFloat(String(r["Value"]).replace(/\D/g,""))||0),0),
      cr=t?((cl/t)*100).toFixed(1):"0.0";
    return[
      {label:"Total Lead",value:t,color:"#185fa5",bg:"#e6f1fb"},
      {label:"Closed",value:cl,color:"#3b6d11",bg:"#eaf3de"},
      {label:"Hot 🔥",value:ht,color:"#a32d2d",bg:"#fcebeb"},
      {label:"Failed",value:fl,color:"#791f1f",bg:"#fce8e8"},
      {label:"Conversion",value:`${cr}%`,color:"#534ab7",bg:"#eeedfe"},
      {label:"Total Value",value:`Rp ${(tv/1e6).toFixed(1)}M`,color:"#854f0b",bg:"#faeeda"},
    ];
  },[leads]);

  const LEAD_VISIBLE=["Lead Id","Date Started","PIC","Platform","Customer Name","Cust Group","Area","Sub Area","Product Group","Value","Status"];
  const PRICE_VISIBLE=["Lead ID","Tanggal","PIC","Product Group","Product Name","Brand","Seller","Area - Sub Area","Price","Year"];

  const ConfirmDelete=({onConfirm,onCancel,label})=>(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:99999,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:12,padding:"28px 32px",width:360,boxShadow:"0 16px 48px rgba(0,0,0,0.18)"}}>
        <div style={{fontWeight:700,fontSize:16,marginBottom:8}}>Hapus {label}?</div>
        <div style={{color:"#666",fontSize:14,marginBottom:24}}>Data akan dihapus permanen dari Firebase.</div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button onClick={onCancel} style={{padding:"9px 18px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#f5f5f5",color:"#444",cursor:"pointer",fontSize:13}}>Batal</button>
          <button onClick={onConfirm} style={{padding:"9px 18px",borderRadius:7,border:"none",background:"#d93025",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>Hapus</button>
        </div>
      </div>
    </div>
  );

  // ── DB Error screen ──
  if(dbError) return (
    <div style={{minHeight:"100vh",background:"#f4f6f9",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px",maxWidth:520,textAlign:"center",boxShadow:"0 8px 32px rgba(0,0,0,0.08)"}}>
        <div style={{fontSize:48,marginBottom:12}}>🔥</div>
        <div style={{fontWeight:700,fontSize:18,marginBottom:12,color:"#d93025"}}>Firebase Belum Terhubung</div>
        <div style={{fontSize:14,color:"#555",marginBottom:24,lineHeight:1.6}}>{dbError}</div>
        <div style={{background:"#f8fafc",borderRadius:8,padding:"16px",textAlign:"left",fontSize:13,color:"#333",fontFamily:"monospace"}}>
          <div style={{marginBottom:8,fontWeight:700,fontFamily:"sans-serif",color:"#666"}}>Langkah setup:</div>
          1. Buka <strong>Firebase Console</strong> → buat project baru<br/>
          2. Aktifkan <strong>Firestore Database</strong><br/>
          3. Buka Project Settings → Your apps → tambah Web app<br/>
          4. Copy config ke <strong>src/firebase.js</strong><br/>
          5. Di Firestore Rules, set: <code style={{background:"#eee",padding:"1px 4px"}}>allow read, write: if true;</code>
        </div>
      </div>
    </div>
  );

  if(!role) return <LoginScreen onLogin={setRole}/>;
  if(showAdmin) return <AdminPanel master={master} setMasterField={setMasterField} onBack={()=>setShowAdmin(false)}/>;

  const isLoading=masterLoading||leadsLoading||pricesLoading;

  return (
    <div style={{minHeight:"100vh",background:"#f4f6f9",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif"}}>
      {toast&&<div style={{position:"fixed",top:16,right:16,zIndex:99999,background:toast.ok?"#eaf3de":"#fcebeb",color:toast.ok?"#2d6a11":"#a32d2d",border:`1px solid ${toast.ok?"#b3d97a":"#f0a0a0"}`,borderRadius:8,padding:"11px 18px",fontSize:13,fontWeight:500,boxShadow:"0 4px 16px rgba(0,0,0,0.10)",pointerEvents:"none"}}>{toast.ok?"✓ ":"✕ "}{toast.msg}</div>}

      {leadDeleteId&&<ConfirmDelete label="lead ini" onConfirm={()=>handleDeleteLead(leadDeleteId)} onCancel={()=>setLeadDeleteId(null)}/>}
      {priceDeleteId&&<ConfirmDelete label="data harga ini" onConfirm={()=>handleDeletePrice(priceDeleteId)} onCancel={()=>setPriceDeleteId(null)}/>}

      {leadModal&&(
        <Modal title={leadModal.type==="add"?"➕ Tambah Lead Baru":"✏️ Edit Lead"} onClose={()=>setLeadModal(null)} wide>
          <LeadForm leads={leads} master={master} initial={leadModal.row||{}} isEdit={leadModal.type==="edit"}
            onSave={(form)=>handleSaveLead(form,leadModal.type==="edit",leadModal.row?._id)}
            onCancel={()=>setLeadModal(null)}/>
        </Modal>
      )}
      {priceModal&&(
        <Modal title={priceModal.type==="add"?"➕ Tambah Data Harga":"✏️ Edit Data Harga"} onClose={()=>setPriceModal(null)} wide>
          <PriceForm master={master} initial={priceModal.row||{}} isEdit={priceModal.type==="edit"}
            onSave={(form)=>handleSavePrice(form,priceModal.type==="edit",priceModal.row?._id)}
            onCancel={()=>setPriceModal(null)}/>
        </Modal>
      )}

      {/* Header */}
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"0 24px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"13px 0 10px"}}>
          <span style={{fontSize:24}}>⚓</span>
          <div>
            <div style={{fontWeight:800,fontSize:16,color:"#1a1a1a"}}>Lead Conversion Dashboard</div>
            <div style={{fontSize:12,color:"#888",display:"flex",alignItems:"center",gap:6}}>
              Marine Engine Sales Tracker
              <span style={{fontSize:11,background:"#eaf3de",color:"#2d6a11",padding:"1px 7px",borderRadius:10,fontWeight:600}}>
                {isLoading?"⏳ Memuat...":"🔥 Firebase Live"}
              </span>
            </div>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontSize:12,background:role==="admin"?"#faeeda":"#e6f1fb",padding:"4px 10px",borderRadius:20,fontWeight:600,color:role==="admin"?"#854f0b":"#185fa5"}}>
              {role==="admin"?"🔐 Admin":"💼 Sales"}
            </span>
            {role==="admin"&&<button onClick={()=>setShowAdmin(true)} style={{padding:"7px 13px",borderRadius:8,border:"1.5px solid #185fa5",background:"#f0f7ff",color:"#185fa5",cursor:"pointer",fontSize:13,fontWeight:600}}>⚙️ Master Data</button>}
            <input ref={importRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{display:"none"}}/>
            {role==="admin"&&<button onClick={()=>importRef.current?.click()} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 13px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#444",cursor:"pointer",fontSize:13,fontWeight:500}}>⬆️ Import</button>}
            <button onClick={handleExport} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 13px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>⬇️ Export</button>
            <button onClick={()=>setRole(null)} style={{padding:"7px 11px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#888",cursor:"pointer",fontSize:12}}>Keluar</button>
          </div>
        </div>
        <div style={{display:"flex",gap:0}}>
          {[{id:"lead",icon:"📊",label:"Lead Tracker",count:leads.length},{id:"price",icon:"💰",label:"Product Price",count:prices.length}].map(t=>(
            <button key={t.id} onClick={()=>setActiveTab(t.id)} style={{padding:"9px 18px",fontSize:13,fontWeight:500,cursor:"pointer",border:"none",background:"transparent",borderBottom:activeTab===t.id?"2.5px solid #185fa5":"2.5px solid transparent",color:activeTab===t.id?"#185fa5":"#777",display:"flex",alignItems:"center",gap:7}}>
              {t.icon} {t.label}
              <span style={{background:activeTab===t.id?"#e6f1fb":"#f0f0f0",color:activeTab===t.id?"#185fa5":"#888",borderRadius:12,padding:"1px 8px",fontSize:11,fontWeight:600}}>{t.count}</span>
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"18px 24px"}}>
        {/* Lead Tracker */}
        {activeTab==="lead"&&(
          <>
            {leads.length>0&&(
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10,marginBottom:16}}>
                {stats.map(s=>(
                  <div key={s.label} style={{background:s.bg,borderRadius:10,padding:"12px 14px",border:`1px solid ${s.color}22`}}>
                    <div style={{fontSize:11,color:s.color,fontWeight:600,marginBottom:4}}>{s.label}</div>
                    <div style={{fontSize:20,fontWeight:800,color:s.color}}>{s.value}</div>
                  </div>
                ))}
              </div>
            )}
            <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
              <input value={leadSearch} onChange={e=>{setLeadSearch(e.target.value);setLeadPage(1);}} placeholder="🔍  Cari lead..."
                style={{flex:"1 1 200px",maxWidth:280,padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none"}}/>
              {statusOptions.length>0&&<select value={filterStatus} onChange={e=>{setFilterStatus(e.target.value);setLeadPage(1);}} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#444",fontSize:13}}>
                <option value="all">Semua Status</option>{statusOptions.map(s=><option key={s} value={s}>{s}</option>)}
              </select>}
              {groupOptions.length>0&&<select value={filterGroup} onChange={e=>{setFilterGroup(e.target.value);setLeadPage(1);}} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#444",fontSize:13}}>
                <option value="all">Semua Grup</option>{groupOptions.map(g=><option key={g} value={g}>{g}</option>)}
              </select>}
              <span style={{fontSize:12,color:"#999"}}>{filteredLeads.length} lead</span>
              <button onClick={()=>setLeadModal({type:"add"})} style={{marginLeft:"auto",padding:"8px 16px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>＋ Lead Baru</button>
            </div>
            {leadsLoading?(
              <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px",textAlign:"center",color:"#aaa",fontSize:14}}>⏳ Memuat data dari Firebase...</div>
            ):leads.length===0?(
              <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px 24px",textAlign:"center"}}>
                <div style={{fontSize:48,marginBottom:12}}>📂</div>
                <div style={{fontWeight:700,fontSize:15,marginBottom:8}}>Belum ada lead</div>
                <div style={{fontSize:13,color:"#888",marginBottom:24}}>Import Excel atau tambah lead baru.</div>
                <div style={{display:"flex",gap:10,justifyContent:"center"}}>
                  {role==="admin"&&<button onClick={()=>importRef.current?.click()} style={{padding:"9px 20px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>⬆️ Import Excel</button>}
                  <button onClick={()=>setLeadModal({type:"add"})} style={{padding:"9px 20px",borderRadius:8,border:"1.5px solid #185fa5",background:"#fff",color:"#185fa5",cursor:"pointer",fontSize:13,fontWeight:600}}>＋ Tambah Lead</button>
                </div>
              </div>
            ):(
              <GenericTable rows={paginatedLeads} visibleCols={LEAD_VISIBLE} page={leadPage} totalPages={leadTotalPages} filtered={filteredLeads}
                onEdit={row=>setLeadModal({type:"edit",row})}
                onDelete={id=>setLeadDeleteId(id)} role={role} onPageChange={setLeadPage}/>
            )}
          </>
        )}

        {/* Product Price */}
        {activeTab==="price"&&(
          <>
            <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
              <input value={priceSearch} onChange={e=>{setPriceSearch(e.target.value);setPricePage(1);}} placeholder="🔍  Cari produk..."
                style={{flex:"1 1 200px",maxWidth:280,padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none"}}/>
              <span style={{fontSize:12,color:"#999"}}>{filteredPrices.length} produk</span>
              <button onClick={()=>setPriceModal({type:"add"})} style={{marginLeft:"auto",padding:"8px 16px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>＋ Tambah Harga</button>
            </div>
            {pricesLoading?(
              <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px",textAlign:"center",color:"#aaa",fontSize:14}}>⏳ Memuat data dari Firebase...</div>
            ):prices.length===0?(
              <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px 24px",textAlign:"center"}}>
                <div style={{fontSize:48,marginBottom:12}}>💰</div>
                <div style={{fontWeight:700,fontSize:15,marginBottom:8}}>Belum ada data harga</div>
                <div style={{display:"flex",gap:10,justifyContent:"center",marginTop:20}}>
                  {role==="admin"&&<button onClick={()=>importRef.current?.click()} style={{padding:"9px 20px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>⬆️ Import Excel</button>}
                  <button onClick={()=>setPriceModal({type:"add"})} style={{padding:"9px 20px",borderRadius:8,border:"1.5px solid #185fa5",background:"#fff",color:"#185fa5",cursor:"pointer",fontSize:13,fontWeight:600}}>＋ Tambah Harga</button>
                </div>
              </div>
            ):(
              <GenericTable rows={paginatedPrices} visibleCols={PRICE_VISIBLE} page={pricePage} totalPages={priceTotalPages} filtered={filteredPrices}
                onEdit={row=>setPriceModal({type:"edit",row})}
                onDelete={id=>setPriceDeleteId(id)} role={role} onPageChange={setPricePage}/>
            )}
          </>
        )}
      </div>
      <style>{`button:disabled{opacity:0.4;cursor:not-allowed!important;} tr:last-child td{border-bottom:none!important;}`}</style>
    </div>
  );
}
