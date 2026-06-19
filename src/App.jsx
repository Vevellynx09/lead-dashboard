import { useState, useRef, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { collection, doc, addDoc, updateDoc, deleteDoc, onSnapshot, setDoc, getDoc } from "firebase/firestore";
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
    { id:1, name:"Jawa Timur", subs:["Surabaya","Malang","Sidoarjo","Gresik","Pasuruan","Tuban","Lamongan"] },
    { id:2, name:"Jawa Tengah", subs:["Semarang","Solo","Yogyakarta","Purwokerto"] },
    { id:3, name:"Bali", subs:["Denpasar","Badung","Gianyar","Buleleng"] },
    { id:4, name:"Kalimantan", subs:["Balikpapan","Samarinda","Banjarmasin","Pontianak"] },
    { id:5, name:"Sulawesi", subs:["Makassar","Manado","Kendari","Palu"] },
    { id:6, name:"Sumatera", subs:["Medan","Palembang","Batam","Pekanbaru"] },
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
const FILTERABLE_COLS = ["PIC","Platform","Cust Group","Area","Sub Area","Product Group","Status","Year","Periode"];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function todayStr() { const d=new Date(); return `${d.getDate().toString().padStart(2,"0")}/${(d.getMonth()+1).toString().padStart(2,"0")}/${d.getFullYear()}`; }
function dateToLeadSuffix(s) { const p=s?s.split("/"):[]; if(p.length!==3){const n=new Date();return`${n.getDate().toString().padStart(2,"0")}${ROMAN[n.getMonth()]}${n.getFullYear()}`;}return`${p[0]}${ROMAN[parseInt(p[1])-1]}${p[2]}`; }
function dateToYear(d) { const p=d?d.split("/"):[]; return p.length===3?p[2]:new Date().getFullYear().toString(); }
function dateToPeriode(d) { const p=d?d.split("/"):[]; if(p.length!==3)return`${INDO_MONTHS[new Date().getMonth()]} ${new Date().getFullYear()}`; return`${INDO_MONTHS[parseInt(p[1])-1]} ${p[2]}`; }
function formatContact(raw) { const d=raw.replace(/\D/g,""); if(d.length<=4)return d; if(d.length<=8)return`${d.slice(0,4)}-${d.slice(4)}`; if(d.length<=12)return`${d.slice(0,4)}-${d.slice(4,8)}-${d.slice(8)}`; return`${d.slice(0,4)}-${d.slice(4,8)}-${d.slice(8,12)}`; }
function normalizePhone(s) { return (s||"").replace(/\D/g,""); }
function genLeadId(leads, dateStr) { const active=leads.filter(l=>!l._deleted); return`${(active.length+1).toString().padStart(3,"0")}-${dateToLeadSuffix(dateStr)}`; }
function fmtValue(v) { const n=parseFloat(String(v).replace(/\D/g,"")); if(!n)return"—"; return`Rp ${n.toLocaleString("id-ID")}`; }
function fmtValueShort(v) { if(!v)return"Rp 0"; if(v>=1e9)return`Rp ${(v/1e9).toFixed(2).replace(".",",")} M`; if(v>=1e6)return`Rp ${(v/1e6).toFixed(2).replace(".",",")} Jt`; return`Rp ${v.toLocaleString("id-ID")}`; }

// ─── Badge ────────────────────────────────────────────────────────────────────
function Badge({ text }) {
  if (!text) return <span style={{color:"#bbb",fontSize:11}}>—</span>;
  const c=STATUS_COLORS[text]||{bg:"#f1f1f1",text:"#555",border:"#ccc"};
  return <span style={{background:c.bg,color:c.text,border:`1px solid ${c.border}`,padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>{text}</span>;
}

// ─── SearchableDropdown ───────────────────────────────────────────────────────
function SearchableDropdown({ value, onChange, options=[], placeholder="Pilih...", disabled=false, error=false }) {
  const [open,setOpen]=useState(false); const [q,setQ]=useState(""); const ref=useRef();
  useEffect(()=>{ const h=e=>{if(ref.current&&!ref.current.contains(e.target)){setOpen(false);setQ("");}}; document.addEventListener("mousedown",h); return()=>document.removeEventListener("mousedown",h); },[]);
  const filtered=options.filter(o=>(typeof o==="string"?o:o.label).toLowerCase().includes(q.toLowerCase()));
  return (
    <div ref={ref} style={{position:"relative"}}>
      <div onClick={()=>!disabled&&setOpen(o=>!o)} style={{padding:"8px 10px",borderRadius:6,border:`1.5px solid ${error?"#d93025":open?"#185fa5":"#d0d5dd"}`,background:disabled?"#f5f5f5":"#fff",color:value?"#1a1a1a":"#aaa",fontSize:13,cursor:disabled?"default":"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",minHeight:36,userSelect:"none"}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{value||placeholder}</span>
        <span style={{marginLeft:8,fontSize:10,color:"#888",flexShrink:0}}>{open?"▲":"▼"}</span>
      </div>
      {open&&(<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",border:"1.5px solid #d0d5dd",borderRadius:8,zIndex:9000,boxShadow:"0 8px 24px rgba(0,0,0,0.14)",overflow:"hidden"}}>
        <input autoFocus value={q} onChange={e=>setQ(e.target.value)} placeholder="Cari..." style={{padding:"8px 12px",border:"none",borderBottom:"1px solid #eee",outline:"none",fontSize:13,width:"100%",boxSizing:"border-box"}}/>
        <div style={{maxHeight:200,overflowY:"auto"}}>
          {filtered.length===0?<div style={{padding:"12px",color:"#aaa",fontSize:13,textAlign:"center"}}>Tidak ditemukan</div>
            :filtered.map((o,i)=>{ const label=typeof o==="string"?o:o.label,val=typeof o==="string"?o:o.value; return(<div key={i} onClick={()=>{onChange(val);setOpen(false);setQ("");}} style={{padding:"9px 12px",cursor:"pointer",fontSize:13,color:"#1a1a1a",borderBottom:"1px solid #f5f5f5"}} onMouseEnter={e=>e.currentTarget.style.background="#f0f7ff"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{label}</div>); })
          }
        </div>
      </div>)}
    </div>
  );
}

// ─── CreatableDropdown ────────────────────────────────────────────────────────
function CreatableDropdown({ value, onChange, options=[], placeholder="Ketik atau pilih...", error=false }) {
  const [open,setOpen]=useState(false); const [q,setQ]=useState(value||""); const ref=useRef();
  useEffect(()=>{setQ(value||"");},[value]);
  useEffect(()=>{ const h=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false);}; document.addEventListener("mousedown",h); return()=>document.removeEventListener("mousedown",h); },[]);
  const filtered=q.trim()?options.filter(o=>o.toLowerCase().includes(q.toLowerCase())):options;
  const isExact=options.some(o=>o.toLowerCase()===q.trim().toLowerCase());
  const isNew=q.trim()&&!isExact;
  return (
    <div ref={ref} style={{position:"relative"}}>
      <div style={{position:"relative"}}>
        <input value={q} onChange={e=>{setQ(e.target.value);onChange(e.target.value);setOpen(true);}} onFocus={()=>setOpen(true)} onBlur={()=>setTimeout(()=>setOpen(false),150)} placeholder={placeholder}
          style={{padding:"8px 10px",paddingRight:q.trim()?90:10,borderRadius:6,border:`1.5px solid ${error?"#d93025":open?"#185fa5":"#d0d5dd"}`,background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}}/>
        {q.trim()&&<span style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",fontSize:10,fontWeight:700,padding:"2px 7px",borderRadius:10,background:isExact?"#eaf3de":"#faeeda",color:isExact?"#2d6a11":"#854f0b",pointerEvents:"none",whiteSpace:"nowrap"}}>{isExact?"✓ Terdaftar":"⊕ Baru"}</span>}
      </div>
      {open&&(filtered.length>0||isNew)&&(<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",border:"1.5px solid #d0d5dd",borderRadius:8,zIndex:9000,boxShadow:"0 8px 24px rgba(0,0,0,0.14)",overflow:"hidden"}}>
        <div style={{maxHeight:180,overflowY:"auto"}}>{filtered.map((opt,i)=>(<div key={i} onMouseDown={()=>{setQ(opt);onChange(opt);setOpen(false);}} style={{padding:"9px 12px",cursor:"pointer",fontSize:13,color:"#1a1a1a",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",gap:8}} onMouseEnter={e=>e.currentTarget.style.background="#f0f7ff"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}><span style={{fontSize:12}}>🏢</span><span style={{flex:1}}>{opt}</span></div>))}</div>
        {isNew&&<div style={{padding:"9px 12px",fontSize:12,color:"#854f0b",background:"#fffbf0",borderTop:filtered.length>0?"1px solid #f5e8c0":"none"}}>⊕ Tambah baru: <strong>"{q.trim()}"</strong></div>}
      </div>)}
    </div>
  );
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function Modal({ title, onClose, children, wide=false }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
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
  const [form,setForm]=useState(()=>{ if(!isEdit){const lid=genLeadId(leads,today);return{"Lead Id":lid,"Date Started":today,"Year":dateToYear(today),"Periode":dateToPeriode(today),"Date Ended":"","PIC":"","Platform":"","Customer Name":"","Parent Name":"","Cust Group":"","Area":"","Sub Area":"","Contact":"","Purpose":"","Product Group":"","Value":"","Status":"1 - Cold","Failed Report":"","Lead Time":"","Notes":""};} return{...initial}; });
  const [errors,setErrors]=useState({});
  const [contactMatch,setContactMatch]=useState(null);
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const parentNameOptions=useMemo(()=>{const names=leads.filter(l=>!l._deleted).map(l=>(l["Parent Name"]||"").trim()).filter(Boolean);return[...new Set(names)].sort();},[leads]);
  const handleContactChange=(raw)=>{ const digits=raw.replace(/\D/g,""); set("Contact",formatContact(digits)); if(digits.length>=10){const match=leads.find(l=>{if(l._deleted)return false;if(isEdit&&l["Lead Id"]===initial["Lead Id"])return false;return normalizePhone(l["Contact"])===digits;});setContactMatch(match?{name:match["Customer Name"]||"",parentName:match["Parent Name"]||""}:null);}else setContactMatch(null); };
  const areaOptions=useMemo(()=>{ const opts=[]; master.areas.forEach(a=>{if(a.subs&&a.subs.length)a.subs.forEach(s=>opts.push({label:`${a.name} › ${s}`,value:`${a.name}||${s}`}));else opts.push({label:a.name,value:`${a.name}||`});}); return opts; },[master.areas]);
  const handleAreaSelect=(val)=>{ const[area,sub]=val.split("||"); setForm(p=>({...p,Area:area,"Sub Area":sub||""})); };
  const areaDisplay=form["Area"]?(form["Sub Area"]?`${form["Area"]} › ${form["Sub Area"]}`:form["Area"]):"";
  const validate=()=>{ const e={}; if(!form["Customer Name"].trim())e["Customer Name"]="Wajib diisi"; if(!form["Contact"].trim())e["Contact"]="Wajib diisi"; else{const d=normalizePhone(form["Contact"]);if(d.length<10||d.length>13)e["Contact"]="Format: 0XXX-XXXX-XXXX";} if(!form["Area"])e["Area"]="Wajib dipilih"; if(!form["Status"])e["Status"]="Wajib dipilih"; setErrors(e);return Object.keys(e).length===0; };
  const handleSave=()=>{ if(!validate())return; const saved={...form}; if(isEdit)saved["Date Ended"]=today; onSave(saved); };
  const field=(label,content,required=false,errKey=null)=>(<div style={{display:"flex",flexDirection:"column",gap:4}}><label style={{fontSize:11,color:"#555",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.04em"}}>{label}{required&&<span style={{color:"#d93025",marginLeft:2}}>*</span>}</label>{content}{errors[errKey||label]&&<span style={{fontSize:11,color:"#d93025"}}>{errors[errKey||label]}</span>}</div>);
  const inp=(k,opts={})=>(<input value={form[k]||""} onChange={e=>set(k,e.target.value)} readOnly={opts.readOnly} placeholder={opts.placeholder||""} style={{padding:"8px 10px",borderRadius:6,border:`1.5px solid ${errors[k]?"#d93025":"#d0d5dd"}`,background:opts.readOnly?"#f5f5f5":"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}} onFocus={e=>{if(!opts.readOnly)e.target.style.borderColor="#185fa5";}} onBlur={e=>{if(!opts.readOnly)e.target.style.borderColor=errors[k]?"#d93025":"#d0d5dd";}}/>);
  const dd=(k,options,placeholder)=>(<SearchableDropdown value={form[k]||""} onChange={v=>set(k,v)} options={options} placeholder={placeholder||"Pilih..."} error={!!errors[k]}/>);
  return (
    <div>
      <div style={{background:"#f0f7ff",borderRadius:8,padding:"10px 14px",marginBottom:16,display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10}}>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Lead ID</span><span style={{fontFamily:"monospace",fontWeight:700,color:"#185fa5",fontSize:13}}>{form["Lead Id"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Tanggal Mulai</span><span style={{fontSize:13,color:"#333"}}>{form["Date Started"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Periode</span><span style={{fontSize:13,color:"#333"}}>{form["Periode"]}</span></div>
        <div><span style={{fontSize:10,color:"#666",fontWeight:700,display:"block",textTransform:"uppercase",letterSpacing:"0.05em"}}>Tgl. Terakhir Edit</span>{isEdit?<span style={{fontSize:12,fontWeight:700,color:"#2d6a11"}}>✓ {today}</span>:<span style={{fontSize:12,color:"#aaa",fontStyle:"italic"}}>Auto saat di-edit</span>}</div>
      </div>
      {contactMatch&&(<div style={{background:"#fffbe6",border:"1.5px solid #f5c100",borderRadius:8,padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:12}}>
        <span style={{fontSize:18}}>👤</span>
        <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700,color:"#7a4f00"}}>Nomor ini sudah terdaftar!</div><div style={{fontSize:12,color:"#555",marginTop:2}}><strong>{contactMatch.name}</strong>{contactMatch.parentName?` · ${contactMatch.parentName}`:""}</div></div>
        <button onClick={()=>{setForm(p=>({...p,"Customer Name":contactMatch.name,"Parent Name":contactMatch.parentName}));setContactMatch(null);}} style={{padding:"6px 14px",borderRadius:6,border:"none",background:"#f5c100",color:"#7a4f00",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>✓ Gunakan Data Ini</button>
        <button onClick={()=>setContactMatch(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#aaa",fontSize:18,lineHeight:1,padding:"0 4px"}}>×</button>
      </div>)}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px 16px",maxHeight:"50vh",overflowY:"auto",paddingRight:4,paddingBottom:4}}>
        {field("Contact",(<input value={form["Contact"]||""} onChange={e=>handleContactChange(e.target.value)} placeholder="08XX-XXXX-XXXX" maxLength={14} style={{padding:"8px 10px",borderRadius:6,border:`1.5px solid ${errors["Contact"]?"#d93025":"#d0d5dd"}`,background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"}} onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor=errors["Contact"]?"#d93025":"#d0d5dd"}/>),true,"Contact")}
        {field("Customer Name",inp("Customer Name",{placeholder:"Nama pelanggan"}),true)}
        {field("Parent Name",(<CreatableDropdown value={form["Parent Name"]||""} onChange={v=>set("Parent Name",v)} options={parentNameOptions} placeholder="Nama perusahaan / merk..."/>))}
        {field("Cust Group",dd("Cust Group",master.custGroup,"Pilih Grup..."))}
        {field("Area (Area › Sub Area)",(<SearchableDropdown value={areaDisplay} onChange={handleAreaSelect} options={areaOptions} placeholder="Pilih Area › Sub Area..." error={!!errors["Area"]}/>),true,"Area")}
        {field("PIC",dd("PIC",master.pic,"Pilih PIC..."))}
        {field("Platform",dd("Platform",master.platform,"Pilih Platform..."))}
        {field("Purpose",dd("Purpose",master.purpose,"Pilih Purpose..."))}
        {field("Product Group",dd("Product Group",master.productGroup,"Pilih Produk..."))}
        {field("Value",(<input value={form["Value"]||""} onChange={e=>set("Value",e.target.value.replace(/\D/g,""))} placeholder="Nilai transaksi (Rp)" style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"}} onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>))}
        {field("Status",dd("Status",STATUS_OPTS,"Pilih Status..."),true)}
        {form["Status"]==="5 - Failed"&&<div style={{gridColumn:"1/-1"}}>{field("Failed Report",inp("Failed Report",{placeholder:"Alasan gagal"}))}</div>}
        <div style={{gridColumn:"1/-1"}}>{field("Notes",<textarea value={form["Notes"]||""} onChange={e=>set("Notes",e.target.value)} placeholder="Catatan tambahan..." style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",resize:"vertical",minHeight:56,fontFamily:"inherit"}} onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>)}</div>
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:14,paddingTop:14,borderTop:"1px solid #e8e8e8",flexShrink:0}}>
        <button onClick={onCancel} style={{padding:"9px 20px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#f5f5f5",color:"#444",cursor:"pointer",fontSize:13,fontWeight:500}}>Batal</button>
        <button onClick={handleSave} style={{padding:"9px 20px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>{isEdit?"💾 Simpan Perubahan":"➕ Tambah Lead"}</button>
      </div>
    </div>
  );
}

// ─── PriceForm ────────────────────────────────────────────────────────────────
function PriceForm({ initial={}, master, onSave, onCancel, isEdit=false }) {
  const today=todayStr();
  const [form,setForm]=useState(()=>isEdit?{...initial}:{"Lead ID":"","Tanggal":today,"PIC":"","Platform":"","Product Group":"","Product Name":"","Brand":"","Seller":"","Area - Sub Area":"","Price":"","Year":dateToYear(today),"Period":dateToPeriode(today),"Notes":""});
  const set=(k,v)=>setForm(p=>({...p,[k]:v}));
  const areaOptions=useMemo(()=>{ const opts=[]; master.areas.forEach(a=>{if(a.subs&&a.subs.length)a.subs.forEach(s=>opts.push({label:`${a.name} › ${s}`,value:`${a.name} › ${s}`}));else opts.push({label:a.name,value:a.name});}); return opts; },[master.areas]);
  const inp=(k,opts={})=>(<input value={form[k]||""} onChange={e=>set(k,e.target.value)} readOnly={opts.readOnly} placeholder={opts.placeholder||""} style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:opts.readOnly?"#f5f5f5":"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"}} onFocus={e=>{if(!opts.readOnly)e.target.style.borderColor="#185fa5";}} onBlur={e=>{if(!opts.readOnly)e.target.style.borderColor="#d0d5dd";}}/>);
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
        {field("Price",(<input value={form["Price"]||""} onChange={e=>set("Price",e.target.value.replace(/\D/g,""))} placeholder="Harga (Rp)" style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box"}} onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>))}
        <div style={{gridColumn:"1/-1"}}>{field("Notes",<textarea value={form["Notes"]||""} onChange={e=>set("Notes",e.target.value)} placeholder="Catatan..." style={{padding:"8px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",resize:"vertical",minHeight:56,fontFamily:"inherit"}} onFocus={e=>e.target.style.borderColor="#185fa5"} onBlur={e=>e.target.style.borderColor="#d0d5dd"}/>)}</div>
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:14,paddingTop:14,borderTop:"1px solid #e8e8e8"}}>
        <button onClick={onCancel} style={{padding:"9px 20px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#f5f5f5",color:"#444",cursor:"pointer",fontSize:13,fontWeight:500}}>Batal</button>
        <button onClick={()=>onSave(form)} style={{padding:"9px 20px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>{isEdit?"💾 Simpan":"➕ Tambah"}</button>
      </div>
    </div>
  );
}

// ─── AdminPanel ────────────────────────────────────────────────────────────────
function AdminPanel({ master, setMasterField, trashedLeads, onRestore, onPermanentDelete, onBack }) {
  const [sec,setSec]=useState("areas");
  const [newAreaName,setNewAreaName]=useState(""); const [newSub,setNewSub]=useState({}); const [newItem,setNewItem]=useState("");
  const [importMsg,setImportMsg]=useState(null); const areaImportRef=useRef();
  const SECS=[{id:"areas",label:"📍 Area & Sub Area"},{id:"pic",label:"👤 PIC / Sales"},{id:"platform",label:"📱 Platform"},{id:"productGroup",label:"⚙️ Product Group"},{id:"purpose",label:"🎯 Purpose"},{id:"custGroup",label:"🏢 Customer Group"},{id:"trash",label:`🗑️ Sampah (${trashedLeads.length})`}];
  const addArea=()=>{ if(!newAreaName.trim())return; setMasterField("areas",[...master.areas,{id:Date.now(),name:newAreaName.trim(),subs:[]}]); setNewAreaName(""); };
  const removeArea=id=>setMasterField("areas",master.areas.filter(a=>a.id!==id));
  const addSub=aId=>{ const s=(newSub[aId]||"").trim(); if(!s)return; setMasterField("areas",master.areas.map(a=>a.id===aId?{...a,subs:[...a.subs,s]}:a)); setNewSub(p=>({...p,[aId]:""})); };
  const removeSub=(aId,s)=>setMasterField("areas",master.areas.map(a=>a.id===aId?{...a,subs:a.subs.filter(x=>x!==s)}:a));
  const addToList=k=>{ if(!newItem.trim()||master[k].includes(newItem.trim()))return; setMasterField(k,[...master[k],newItem.trim()]); setNewItem(""); };
  const removeFromList=(k,item)=>setMasterField(k,master[k].filter(i=>i!==item));
  const downloadAreaTemplate=()=>{ const data=[["Area","Sub Area"],["Jawa Timur","Surabaya"],["Jawa Timur","Malang"],["Bali","Denpasar"]]; const ws=XLSX.utils.aoa_to_sheet(data); ws["!cols"]=[{wch:20},{wch:20}]; const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Area Template"); XLSX.writeFile(wb,"Template_Area_SubArea.xlsx"); };
  const handleAreaImport=e=>{ const file=e.target.files[0]; if(!file)return; const reader=new FileReader(); reader.onload=evt=>{ try{ const wb=XLSX.read(evt.target.result,{type:"binary"}); const ws=wb.Sheets[wb.SheetNames[0]]; const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:false}).filter(r=>r.some(c=>c)); const start=rows.length>0&&String(rows[0][0]||"").toLowerCase()==="area"?1:0; const areaMap={}; rows.slice(start).forEach(row=>{const area=String(row[0]||"").trim(),sub=String(row[1]||"").trim(); if(!area)return; if(!areaMap[area])areaMap[area]=new Set(); if(sub)areaMap[area].add(sub);}); const existing=[...master.areas]; Object.entries(areaMap).forEach(([areaName,subsSet])=>{const idx=existing.findIndex(a=>a.name.toLowerCase()===areaName.toLowerCase()); if(idx>=0){existing[idx]={...existing[idx],subs:[...new Set([...existing[idx].subs,...subsSet])]};}else{existing.push({id:Date.now()+Math.random(),name:areaName,subs:[...subsSet]});}}); setMasterField("areas",existing); const ta=Object.keys(areaMap).length,ts=Object.values(areaMap).reduce((s,v)=>s+v.size,0); setImportMsg({ok:true,text:`Berhasil import ${ta} area dan ${ts} sub area`}); setTimeout(()=>setImportMsg(null),4000); }catch(err){setImportMsg({ok:false,text:"Gagal baca file: "+err.message});setTimeout(()=>setImportMsg(null),4000);} }; reader.readAsBinaryString(file); e.target.value=""; };
  return (
    <div style={{minHeight:"100vh",background:"#f4f6f9",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif"}}>
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"14px 24px",display:"flex",alignItems:"center",gap:12}}>
        <button onClick={onBack} style={{background:"#f0f7ff",border:"1px solid #b5d4f4",borderRadius:7,padding:"7px 14px",cursor:"pointer",color:"#185fa5",fontSize:13,fontWeight:600}}>← Dashboard</button>
        <div><div style={{fontWeight:700,fontSize:16,color:"#1a1a1a"}}>⚙️ Admin — Master Data</div><div style={{fontSize:12,color:"#888"}}>Perubahan tersimpan otomatis ke Firebase</div></div>
        <span style={{marginLeft:"auto",fontSize:12,background:"#eaf3de",color:"#2d6a11",padding:"4px 10px",borderRadius:20,fontWeight:600}}>🔥 Firebase Connected</span>
      </div>
      <div style={{display:"flex",gap:0,padding:"20px 24px"}}>
        <div style={{width:200,flexShrink:0,marginRight:18}}>
          {SECS.map(s=>(<button key={s.id} onClick={()=>setSec(s.id)} style={{display:"block",width:"100%",textAlign:"left",padding:"10px 14px",borderRadius:8,border:"none",background:sec===s.id?(s.id==="trash"?"#fce8e8":"#185fa5"):"transparent",color:sec===s.id?(s.id==="trash"?"#d93025":"#fff"):"#444",fontSize:13,fontWeight:sec===s.id?600:400,cursor:"pointer",marginBottom:4}}>{s.label}</button>))}
        </div>
        <div style={{flex:1,background:"#fff",borderRadius:12,border:"1px solid #e5e7eb",padding:20}}>
          {sec==="trash"?(
            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>🗑️ Sampah — Lead yang Dihapus</div>
              <div style={{fontSize:12,color:"#888",marginBottom:16}}>Lead di sini menunggu persetujuan Admin. Restore untuk kembalikan, atau hapus permanen.</div>
              {trashedLeads.length===0?(<div style={{textAlign:"center",padding:"40px 0",color:"#bbb",fontSize:13}}>Tidak ada lead di sampah</div>):(
                <div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {trashedLeads.map(lead=>(<div key={lead._id} style={{border:"1px solid #e5e7eb",borderRadius:10,padding:"12px 16px",background:"#fafbfc",display:"flex",alignItems:"center",gap:12}}>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                        <span style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#185fa5"}}>{lead["Lead Id"]}</span>
                        <Badge text={lead["Status"]}/>
                      </div>
                      <div style={{fontSize:13,fontWeight:600,color:"#1a1a1a"}}>{lead["Customer Name"]||"—"}</div>
                      <div style={{fontSize:11,color:"#888",marginTop:2}}>PIC: {lead["PIC"]||"—"} · {lead["Area"]||"—"} · Dihapus: {lead["_deletedAt"]||"—"}</div>
                    </div>
                    <div style={{display:"flex",gap:6,flexShrink:0}}>
                      <button onClick={()=>onRestore(lead._id)} style={{padding:"7px 14px",borderRadius:7,border:"none",background:"#eaf3de",color:"#2d6a11",cursor:"pointer",fontSize:12,fontWeight:600}}>↩️ Restore</button>
                      <button onClick={()=>onPermanentDelete(lead._id)} style={{padding:"7px 14px",borderRadius:7,border:"none",background:"#fce8e8",color:"#d93025",cursor:"pointer",fontSize:12,fontWeight:600}}>🗑️ Hapus Permanen</button>
                    </div>
                  </div>))}
                </div>
              )}
            </div>
          ):sec==="areas"?(
            <div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{fontWeight:700,fontSize:15}}>📍 Area & Sub Area</div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={downloadAreaTemplate} style={{padding:"7px 13px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#fff",color:"#555",cursor:"pointer",fontSize:12,fontWeight:500}}>📄 Unduh Template</button>
                  <input ref={areaImportRef} type="file" accept=".xlsx,.xls" onChange={handleAreaImport} style={{display:"none"}}/>
                  <button onClick={()=>areaImportRef.current?.click()} style={{padding:"7px 13px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600}}>⬆️ Import Excel</button>
                </div>
              </div>
              <div style={{background:"#fffbf0",border:"1px solid #f5e8c0",borderRadius:8,padding:"9px 13px",marginBottom:14,fontSize:12,color:"#7a4f00"}}>💡 <strong>Format:</strong> Kolom A = Area, Kolom B = Sub Area. Satu baris per kombinasi.</div>
              {importMsg&&<div style={{background:importMsg.ok?"#eaf3de":"#fce8e8",border:`1px solid ${importMsg.ok?"#b3d97a":"#f0a0a0"}`,borderRadius:8,padding:"9px 13px",marginBottom:14,fontSize:13,fontWeight:500,color:importMsg.ok?"#2d6a11":"#a32d2d"}}>{importMsg.ok?"✓":"✕"} {importMsg.text}</div>}
              <div style={{display:"flex",gap:8,marginBottom:20}}>
                <input value={newAreaName} onChange={e=>setNewAreaName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addArea()} placeholder="Nama Area baru..." style={{flex:1,padding:"8px 12px",borderRadius:7,border:"1.5px solid #d0d5dd",fontSize:13,outline:"none"}}/>
                <button onClick={addArea} style={{padding:"8px 16px",borderRadius:7,border:"none",background:"#eaf3de",color:"#2d6a11",cursor:"pointer",fontSize:13,fontWeight:600}}>+ Tambah Area</button>
              </div>
              {master.areas.map(area=>(<div key={area.id} style={{border:"1px solid #e5e7eb",borderRadius:10,marginBottom:12,overflow:"hidden"}}>
                <div style={{background:"#f8fafc",padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontWeight:600,fontSize:14}}>🗺️ {area.name}</span>
                  <button onClick={()=>removeArea(area.id)} style={{background:"#fce8e8",border:"none",color:"#d93025",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:600}}>Hapus</button>
                </div>
                <div style={{padding:"12px 14px"}}>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>
                    {area.subs.map(s=>(<span key={s} style={{background:"#e6f1fb",color:"#185fa5",borderRadius:20,padding:"3px 10px",fontSize:12,fontWeight:500,display:"flex",alignItems:"center",gap:4}}>{s}<button onClick={()=>removeSub(area.id,s)} style={{background:"none",border:"none",cursor:"pointer",color:"#a32d2d",fontSize:14,lineHeight:1,padding:0}}>×</button></span>))}
                    {area.subs.length===0&&<span style={{color:"#bbb",fontSize:12}}>Belum ada sub area</span>}
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <input value={newSub[area.id]||""} onChange={e=>setNewSub(p=>({...p,[area.id]:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&addSub(area.id)} placeholder="Sub Area baru..." style={{flex:1,padding:"7px 10px",borderRadius:6,border:"1.5px solid #d0d5dd",fontSize:13,outline:"none"}}/>
                    <button onClick={()=>addSub(area.id)} style={{padding:"7px 14px",borderRadius:6,border:"none",background:"#eaf3de",color:"#3b6d11",cursor:"pointer",fontSize:13,fontWeight:600}}>+ Sub Area</button>
                  </div>
                </div>
              </div>))}
            </div>
          ):(
            <div>
              <div style={{fontWeight:700,fontSize:15,marginBottom:16}}>{SECS.find(s=>s.id===sec)?.label}</div>
              <div style={{display:"flex",gap:8,marginBottom:16}}>
                <input value={newItem} onChange={e=>setNewItem(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addToList(sec)} placeholder="Tambah item baru..." style={{flex:1,padding:"8px 12px",borderRadius:7,border:"1.5px solid #d0d5dd",fontSize:13,outline:"none"}}/>
                <button onClick={()=>addToList(sec)} style={{padding:"8px 16px",borderRadius:7,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>+ Tambah</button>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {master[sec].map((item,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 14px",border:"1px solid #e5e7eb",borderRadius:8,background:"#fafbfc"}}>
                  <span style={{fontSize:13}}>{item}</span>
                  <button onClick={()=>removeFromList(sec,item)} style={{background:"#fce8e8",border:"none",color:"#d93025",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:600}}>Hapus</button>
                </div>))}
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
        <div style={{fontSize:12,background:"#eaf3de",color:"#2d6a11",padding:"4px 12px",borderRadius:20,display:"inline-block",marginBottom:28,fontWeight:600}}>🔥 Powered by Firebase</div>
        <div style={{fontWeight:600,fontSize:14,color:"#555",marginBottom:16,textAlign:"left"}}>Masuk sebagai:</div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={()=>onLogin("admin")} style={{padding:"14px 20px",borderRadius:10,border:"none",background:"linear-gradient(135deg,#185fa5,#1e7fc4)",color:"#fff",cursor:"pointer",fontSize:15,fontWeight:700,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:24}}>🔐</span><div style={{textAlign:"left"}}><div>Admin</div><div style={{fontSize:11,fontWeight:400,opacity:0.85}}>Kelola master data & semua lead</div></div>
          </button>
          <button onClick={()=>onLogin("sales")} style={{padding:"14px 20px",borderRadius:10,border:"2px solid #185fa5",background:"#f0f7ff",color:"#185fa5",cursor:"pointer",fontSize:15,fontWeight:700,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:24}}>💼</span><div style={{textAlign:"left"}}><div>Sales</div><div style={{fontSize:11,fontWeight:400,opacity:0.8}}>Input & kelola lead pelanggan</div></div>
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── LeadTable dengan sort + column filter ────────────────────────────────────
function LeadTable({ rows, allRows, page, totalPages, filtered, onEdit, onDelete, role, onPageChange, sortCol, sortDir, onSort, colFilters, onColFilter }) {
  const COLS = ["Lead Id","Date Started","PIC","Platform","Customer Name","Cust Group","Area","Sub Area","Product Group","Value","Status"];
  const [openFilter,setOpenFilter]=useState(null);
  const dropdownRef=useRef();

  useEffect(()=>{
    const h=e=>{
      if(dropdownRef.current&&!dropdownRef.current.contains(e.target)) setOpenFilter(null);
    };
    document.addEventListener("mousedown",h);
    return()=>document.removeEventListener("mousedown",h);
  },[]);

  const colOptions=(col)=>[...new Set(allRows.map(r=>r[col]).filter(Boolean))].sort();

  return (
    <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,overflow:"hidden"}}>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5}}>
          <thead>
            <tr style={{background:"#f8fafc",borderBottom:"1px solid #e5e7eb"}}>
              <th style={{padding:"10px 10px",textAlign:"center",color:"#aaa",fontWeight:600,fontSize:10,width:36}}>#</th>
              {COLS.map(col=>{
                const isSorted=sortCol===col;
                const isFiltered=!!colFilters[col];
                const canFilter=FILTERABLE_COLS.includes(col);
                const isOpen=openFilter===col;
                return (
                  <th key={col} style={{padding:"8px 10px",textAlign:"left",color:isSorted||isFiltered?"#185fa5":"#555",fontWeight:700,fontSize:10,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.04em",background:isFiltered?"#eef6ff":"transparent",position:"relative",userSelect:"none"}}>
                    <div style={{display:"flex",alignItems:"center",gap:5}}>
                      {/* Sort trigger */}
                      <span onClick={()=>onSort(col)} style={{cursor:"pointer",display:"flex",alignItems:"center",gap:3,flex:1}}>
                        {col}
                        <span style={{fontSize:9,color:isSorted?"#185fa5":"#ccc",marginLeft:1}}>{isSorted?(sortDir==="asc"?"▲":"▼"):"⇅"}</span>
                      </span>

                      {/* Filter button */}
                      {canFilter&&(
                        <div style={{position:"relative"}} ref={isOpen?dropdownRef:null}>
                          <button
                            onClick={(e)=>{e.stopPropagation();setOpenFilter(isOpen?null:col);}}
                            title={`Filter ${col}`}
                            style={{
                              background:isFiltered?"#185fa5":isOpen?"#d0e4f7":"#e8edf2",
                              border:"none",
                              borderRadius:5,
                              width:22,height:22,
                              cursor:"pointer",
                              fontSize:10,
                              color:isFiltered||isOpen?"#fff":"#555",
                              display:"flex",alignItems:"center",justifyContent:"center",
                              padding:0,flexShrink:0,
                              fontWeight:700,
                            }}>
                            {isFiltered?"✕":"▼"}
                          </button>

                          {isOpen&&(
                            <div style={{
                              position:"absolute",top:"calc(100% + 6px)",left:"50%",transform:"translateX(-50%)",
                              background:"#fff",border:"1.5px solid #d0d5dd",borderRadius:9,
                              zIndex:8000,boxShadow:"0 8px 28px rgba(0,0,0,0.16)",
                              minWidth:170,maxHeight:240,overflow:"auto",
                            }}>
                              <div style={{padding:"8px 12px",fontSize:11,color:"#888",borderBottom:"1px solid #f0f0f0",fontWeight:600,background:"#f9fafb",borderRadius:"7px 7px 0 0"}}>
                                Filter: {col}
                              </div>
                              <div
                                onClick={()=>{onColFilter(col,"");setOpenFilter(null);}}
                                style={{padding:"8px 12px",cursor:"pointer",fontSize:12,color:!colFilters[col]?"#185fa5":"#888",fontStyle:"italic",borderBottom:"1px solid #f5f5f5",fontWeight:!colFilters[col]?700:400,background:!colFilters[col]?"#f0f7ff":"transparent"}}
                                onMouseEnter={e=>{if(colFilters[col])e.currentTarget.style.background="#f5f5f5";}}
                                onMouseLeave={e=>{if(colFilters[col])e.currentTarget.style.background="transparent";}}>
                                Semua {col}
                              </div>
                              {colOptions(col).map(opt=>(
                                <div key={opt}
                                  onClick={()=>{onColFilter(col,opt);setOpenFilter(null);}}
                                  style={{padding:"8px 12px",cursor:"pointer",fontSize:12,color:colFilters[col]===opt?"#185fa5":"#1a1a1a",fontWeight:colFilters[col]===opt?700:400,background:colFilters[col]===opt?"#e6f1fb":"transparent",borderBottom:"1px solid #f5f5f5",display:"flex",alignItems:"center",gap:6}}
                                  onMouseEnter={e=>{if(colFilters[col]!==opt)e.currentTarget.style.background="#f0f7ff";}}
                                  onMouseLeave={e=>{if(colFilters[col]!==opt)e.currentTarget.style.background=colFilters[col]===opt?"#e6f1fb":"transparent";}}>
                                  {colFilters[col]===opt&&<span style={{fontSize:10,color:"#185fa5"}}>✓</span>}
                                  <span>{opt}</span>
                                </div>
                              ))}
                              {colOptions(col).length===0&&<div style={{padding:"12px",color:"#bbb",fontSize:12,textAlign:"center"}}>Tidak ada data</div>}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  </th>
                );
              })}
              <th style={{padding:"10px 12px",color:"#aaa",fontSize:10,width:80,textAlign:"center",textTransform:"uppercase",letterSpacing:"0.05em"}}>Aksi</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row,i)=>{
              const globalIdx=(page-1)*PAGE_SIZE+i;
              return (
                <tr key={row._id||i} style={{borderBottom:"1px solid #f0f0f0"}} onMouseEnter={e=>e.currentTarget.style.background="#fafbfc"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={{padding:"8px 10px",textAlign:"center",color:"#bbb",fontSize:11}}>{globalIdx+1}</td>
                  {COLS.map(col=>(
                    <td key={col} style={{padding:"8px 10px",color:"#1a1a1a",maxWidth:160}}>
                      {col==="Status"?<Badge text={row[col]}/>
                       :col==="Lead Id"?<span style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#185fa5"}}>{row[col]}</span>
                       :col==="Value"?<span style={{fontWeight:700,color:row[col]?"#2d6a11":"#bbb",fontSize:11}}>{fmtValue(row[col])}</span>
                       :col==="Cust Group"?<span style={{fontSize:11,fontWeight:500,color:"#534ab7"}}>{row[col]||"—"}</span>
                       :<span style={{display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:140,fontSize:12}} title={String(row[col]||"")}>{row[col]||<span style={{color:"#ccc"}}>—</span>}</span>
                      }
                    </td>
                  ))}
                  <td style={{padding:"8px 10px",textAlign:"center",whiteSpace:"nowrap"}}>
                    <button onClick={()=>onEdit(row)} style={{background:"#e8f0fb",border:"none",cursor:"pointer",color:"#185fa5",padding:"5px 8px",borderRadius:6,fontSize:13,marginRight:4}}>✏️</button>
                    <button onClick={()=>onDelete(row._id)} style={{background:"#fce8e8",border:"none",cursor:"pointer",color:"#a32d2d",padding:"5px 8px",borderRadius:6,fontSize:13}}>🗑️</button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {totalPages>1&&(<div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:6,padding:"12px 16px",borderTop:"1px solid #f0f0f0"}}>
        <button onClick={()=>onPageChange(p=>Math.max(1,p-1))} disabled={page===1} style={{padding:"6px 12px",borderRadius:6,border:"1px solid #e0e0e0",background:"#fff",cursor:"pointer",fontSize:12,color:"#555"}}>‹</button>
        {Array.from({length:Math.min(7,totalPages)},(_,i)=>{const s=Math.max(1,Math.min(page-3,totalPages-6));return s+i;}).filter(p=>p>=1&&p<=totalPages).map(p2=>(<button key={p2} onClick={()=>onPageChange(p2)} style={{padding:"6px 10px",minWidth:32,borderRadius:6,border:page===p2?"none":"1px solid #e0e0e0",background:page===p2?"#185fa5":"#fff",color:page===p2?"#fff":"#555",cursor:"pointer",fontSize:12,fontWeight:page===p2?700:400}}>{p2}</button>))}
        <button onClick={()=>onPageChange(p=>Math.min(totalPages,p+1))} disabled={page===totalPages} style={{padding:"6px 12px",borderRadius:6,border:"1px solid #e0e0e0",background:"#fff",cursor:"pointer",fontSize:12,color:"#555"}}>›</button>
        <span style={{fontSize:12,color:"#aaa",marginLeft:4}}>{(page-1)*PAGE_SIZE+1}–{Math.min(page*PAGE_SIZE,filtered.length)} / {filtered.length}</span>
      </div>)}
    </div>
  );
}

// ─── PriceTable ────────────────────────────────────────────────────────────────
function PriceTable({ rows, page, totalPages, filtered, onEdit, onDelete, role, onPageChange }) {
  const COLS=["Lead ID","Tanggal","PIC","Product Group","Product Name","Brand","Seller","Area - Sub Area","Price","Year"];
  return (
    <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,overflow:"hidden"}}>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5}}>
          <thead><tr style={{background:"#f8fafc",borderBottom:"1px solid #e5e7eb"}}>
            <th style={{padding:"10px 10px",textAlign:"center",color:"#aaa",fontWeight:600,fontSize:10,width:36}}>#</th>
            {COLS.map(col=>(<th key={col} style={{padding:"10px 12px",textAlign:"left",color:"#666",fontWeight:700,fontSize:10,whiteSpace:"nowrap",textTransform:"uppercase",letterSpacing:"0.05em"}}>{col}</th>))}
            <th style={{padding:"10px 12px",color:"#aaa",fontSize:10,width:80,textAlign:"center",textTransform:"uppercase"}}>Aksi</th>
          </tr></thead>
          <tbody>{rows.map((row,i)=>{
            const globalIdx=(page-1)*PAGE_SIZE+i;
            return (<tr key={row._id||i} style={{borderBottom:"1px solid #f0f0f0"}} onMouseEnter={e=>e.currentTarget.style.background="#fafbfc"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
              <td style={{padding:"8px 10px",textAlign:"center",color:"#bbb",fontSize:11}}>{globalIdx+1}</td>
              {COLS.map(col=>(<td key={col} style={{padding:"8px 12px",color:"#1a1a1a",maxWidth:160}}>
                {col==="Lead ID"?<span style={{fontFamily:"monospace",fontSize:11,fontWeight:700,color:"#185fa5"}}>{row[col]}</span>
                 :col==="Price"?<span style={{fontWeight:700,color:row[col]?"#2d6a11":"#bbb",fontSize:11}}>{fmtValue(row[col])}</span>
                 :<span style={{display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:150,fontSize:12}} title={String(row[col]||"")}>{row[col]||<span style={{color:"#ccc"}}>—</span>}</span>}
              </td>))}
              <td style={{padding:"8px 10px",textAlign:"center",whiteSpace:"nowrap"}}>
                <button onClick={()=>onEdit(row)} style={{background:"#e8f0fb",border:"none",cursor:"pointer",color:"#185fa5",padding:"5px 8px",borderRadius:6,fontSize:13,marginRight:4}}>✏️</button>
                {role==="admin"&&<button onClick={()=>onDelete(row._id)} style={{background:"#fce8e8",border:"none",cursor:"pointer",color:"#a32d2d",padding:"5px 8px",borderRadius:6,fontSize:13}}>🗑️</button>}
              </td>
            </tr>);
          })}</tbody>
        </table>
      </div>
      {totalPages>1&&(<div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:6,padding:"12px 16px",borderTop:"1px solid #f0f0f0"}}>
        <button onClick={()=>onPageChange(p=>Math.max(1,p-1))} disabled={page===1} style={{padding:"6px 12px",borderRadius:6,border:"1px solid #e0e0e0",background:"#fff",cursor:"pointer",fontSize:12,color:"#555"}}>‹</button>
        {Array.from({length:Math.min(7,totalPages)},(_,i)=>{const s=Math.max(1,Math.min(page-3,totalPages-6));return s+i;}).filter(p=>p>=1&&p<=totalPages).map(p2=>(<button key={p2} onClick={()=>onPageChange(p2)} style={{padding:"6px 10px",minWidth:32,borderRadius:6,border:page===p2?"none":"1px solid #e0e0e0",background:page===p2?"#185fa5":"#fff",color:page===p2?"#fff":"#555",cursor:"pointer",fontSize:12,fontWeight:page===p2?700:400}}>{p2}</button>))}
        <button onClick={()=>onPageChange(p=>Math.min(totalPages,p+1))} disabled={page===totalPages} style={{padding:"6px 12px",borderRadius:6,border:"1px solid #e0e0e0",background:"#fff",cursor:"pointer",fontSize:12,color:"#555"}}>›</button>
        <span style={{fontSize:12,color:"#aaa",marginLeft:4}}>{(page-1)*PAGE_SIZE+1}–{Math.min(page*PAGE_SIZE,filtered.length)} / {filtered.length}</span>
      </div>)}
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
  const [allLeads,setAllLeads]=useState([]);
  const [leadsLoading,setLeadsLoading]=useState(true);
  const [leadSearch,setLeadSearch]=useState("");
  const [filterStatus,setFilterStatus]=useState("all");
  const [filterGroup,setFilterGroup]=useState("all");
  const [statFilter,setStatFilter]=useState(null);
  const [sortCol,setSortCol]=useState("Date Started");
  const [sortDir,setSortDir]=useState("desc");
  const [colFilters,setColFilters]=useState({});
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

  const leads=useMemo(()=>allLeads.filter(l=>!l._deleted),[allLeads]);
  const trashedLeads=useMemo(()=>allLeads.filter(l=>l._deleted),[allLeads]);

  useEffect(()=>{ const ref=doc(db,"config","masterData"); getDoc(ref).then(snap=>{ if(snap.exists()){const data=snap.data();setMaster(prev=>({...DEFAULT_MASTER,...prev,...data}));}else{setDoc(ref,DEFAULT_MASTER).catch(()=>{});} setMasterLoading(false); }).catch(err=>{setDbError("Tidak dapat terhubung ke Firebase. Periksa konfigurasi di firebase.js");setMasterLoading(false);}); },[]);
  useEffect(()=>{ const unsub=onSnapshot(collection(db,"leads"),(snap)=>{ const data=snap.docs.map(d=>({_id:d.id,...d.data()})); setAllLeads(data); setLeadsLoading(false); },(err)=>setLeadsLoading(false)); return()=>unsub(); },[]);
  useEffect(()=>{ const unsub=onSnapshot(collection(db,"prices"),(snap)=>{ const data=snap.docs.map(d=>({_id:d.id,...d.data()})); data.sort((a,b)=>(b["Tanggal"]||"").localeCompare(a["Tanggal"]||"")); setPrices(data); setPricesLoading(false); },(err)=>setPricesLoading(false)); return()=>unsub(); },[]);

  const setMasterField=useCallback((field,value)=>{ const updated={...master,[field]:value}; setMaster(updated); setDoc(doc(db,"config","masterData"),updated).catch(err=>showToast("Gagal simpan ke Firebase: "+err.message,false)); },[master]);

  const handleImport=useCallback(e=>{ const file=e.target.files[0]; if(!file)return; const reader=new FileReader(); reader.onload=async evt=>{ try{ const wb=XLSX.read(evt.target.result,{type:"binary",cellDates:true}); let imported=0;
    if(wb.Sheets["Lead Tracker"]){const raw=XLSX.utils.sheet_to_json(wb.Sheets["Lead Tracker"],{header:1,raw:false}); const headerRowIdx=raw.findIndex(row=>row.some(cell=>["Customer Name","Lead Id","Contact"].includes(String(cell||"").trim()))); if(headerRowIdx<0){showToast("Header kolom tidak ditemukan di sheet Lead Tracker!",false);return;} const hdr=raw[headerRowIdx]; const custNameIdx=hdr.findIndex(h=>String(h||"").trim()==="Customer Name"); const contactIdx=hdr.findIndex(h=>String(h||"").trim()==="Contact"); const rows=raw.slice(headerRowIdx+1).filter(r=>{ if(!r||r.length===0)return false; const custName=custNameIdx>=0?String(r[custNameIdx]||"").trim():""; const contact=contactIdx>=0?String(r[contactIdx]||"").trim():""; return custName!==""||contact!==""; }).map(row=>{const obj={};LEAD_COLS.forEach(col=>{const idx=hdr.findIndex(h=>String(h||"").trim()===col);obj[col]=idx>=0?(row[idx]??""):"";});return obj;}); for(const row of rows){await addDoc(collection(db,"leads"),row);imported++;}}
    if(wb.Sheets["Product Price"]){const raw=XLSX.utils.sheet_to_json(wb.Sheets["Product Price"],{header:1,raw:false}); const headerRowIdx=raw.findIndex(row=>row.some(cell=>["Product Name","Product Group","Lead ID"].includes(String(cell||"").trim()))); if(headerRowIdx>=0){const hdr=raw[headerRowIdx]; const prodNameIdx=hdr.findIndex(h=>String(h||"").trim()==="Product Name"); const prodGrpIdx=hdr.findIndex(h=>String(h||"").trim()==="Product Group"); const rows=raw.slice(headerRowIdx+1).filter(r=>{if(!r||r.length===0)return false;const p1=prodNameIdx>=0?String(r[prodNameIdx]||"").trim():"";const p2=prodGrpIdx>=0?String(r[prodGrpIdx]||"").trim():"";return p1!==""||p2!=="";}).map(row=>{const obj={};PRICE_COLS.forEach(col=>{const idx=hdr.findIndex(h=>String(h||"").trim()===col);obj[col]=idx>=0?(row[idx]??""):"";});return obj;}); for(const row of rows){await addDoc(collection(db,"prices"),row);imported++;}}}
    showToast(`${imported} baris berhasil diimpor ke Firebase`); }catch(err){showToast("Gagal import: "+err.message,false);}};  reader.readAsBinaryString(file); e.target.value=""; },[]);

  const handleExport=useCallback(()=>{ if(!leads.length&&!prices.length){showToast("Tidak ada data!",false);return;} const clean=arr=>arr.map(({_id,_deleted,_deletedAt,...rest})=>rest); const wb=XLSX.utils.book_new(); if(leads.length){const ws=XLSX.utils.json_to_sheet(clean(leads),{header:LEAD_COLS});XLSX.utils.book_append_sheet(wb,ws,"Lead Tracker");} if(prices.length){const ws=XLSX.utils.json_to_sheet(clean(prices),{header:PRICE_COLS});XLSX.utils.book_append_sheet(wb,ws,"Product Price");} XLSX.writeFile(wb,"LeadConversion_export.xlsx"); showToast("File berhasil didownload!"); },[leads,prices]);

  // ── Sort handler ──
  const handleSort=useCallback(col=>{ if(sortCol===col){setSortDir(d=>d==="asc"?"desc":"asc");}else{setSortCol(col);setSortDir("asc");} setLeadPage(1); },[sortCol]);
  const handleColFilter=useCallback((col,val)=>{ setColFilters(p=>({...p,[col]:val})); setLeadPage(1); },[]);

  // ── Lead CRUD (soft delete) ──
  const handleSaveLead=async(form,isEdit,docId)=>{ const{_id,_deleted,_deletedAt,...data}=form; try{if(isEdit){await updateDoc(doc(db,"leads",docId),data);showToast("Lead diperbarui!");}else{await addDoc(collection(db,"leads"),data);showToast("Lead baru ditambahkan!");} setLeadModal(null);}catch(err){showToast("Gagal simpan: "+err.message,false);} };
  const handleDeleteLead=async id=>{ try{await updateDoc(doc(db,"leads",id),{_deleted:true,_deletedAt:todayStr()});setLeadDeleteId(null);showToast("Lead dipindahkan ke sampah 🗑️");}catch(err){showToast("Gagal: "+err.message,false);} };
  const handleRestoreLead=async id=>{ try{await updateDoc(doc(db,"leads",id),{_deleted:false,_deletedAt:null});showToast("Lead berhasil dipulihkan! ✅");}catch(err){showToast("Gagal: "+err.message,false);} };
  const handlePermanentDelete=async id=>{ try{await deleteDoc(doc(db,"leads",id));showToast("Lead dihapus permanen.");}catch(err){showToast("Gagal: "+err.message,false);} };

  // ── Price CRUD ──
  const handleSavePrice=async(form,isEdit,docId)=>{ const{_id,...data}=form; try{if(isEdit){await updateDoc(doc(db,"prices",docId),data);showToast("Data harga diperbarui!");}else{await addDoc(collection(db,"prices"),data);showToast("Data harga ditambahkan!");} setPriceModal(null);}catch(err){showToast("Gagal simpan: "+err.message,false);} };
  const handleDeletePrice=async id=>{ try{await deleteDoc(doc(db,"prices",id));setPriceDeleteId(null);showToast("Data harga dihapus.");}catch(err){showToast("Gagal hapus: "+err.message,false);} };

  const statusOptions=useMemo(()=>[...new Set(leads.map(r=>r["Status"]).filter(Boolean))].sort(),[leads]);
  const groupOptions=useMemo(()=>[...new Set(leads.map(r=>r["Cust Group"]).filter(Boolean))].sort(),[leads]);

  const filteredLeads=useMemo(()=>{
    let d=leads;
    if(leadSearch.trim()){const q=leadSearch.toLowerCase();d=d.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q)));}
    if(filterStatus!=="all") d=d.filter(r=>r["Status"]===filterStatus);
    if(filterGroup!=="all") d=d.filter(r=>r["Cust Group"]===filterGroup);
    if(statFilter) d=d.filter(r=>r["Status"]===statFilter);
    Object.entries(colFilters).forEach(([col,val])=>{if(val)d=d.filter(r=>r[col]===val);});
    if(sortCol){d=[...d].sort((a,b)=>{const isNum=sortCol==="Value"||sortCol==="Lead Time"; if(isNum){const na=parseFloat(String(a[sortCol]||"0").replace(/\D/g,""))||0;const nb=parseFloat(String(b[sortCol]||"0").replace(/\D/g,""))||0;return sortDir==="asc"?na-nb:nb-na;}const va=String(a[sortCol]||"").toLowerCase();const vb=String(b[sortCol]||"").toLowerCase();return sortDir==="asc"?va.localeCompare(vb):vb.localeCompare(va);});}
    return d;
  },[leads,leadSearch,filterStatus,filterGroup,statFilter,colFilters,sortCol,sortDir]);

  const filteredPrices=useMemo(()=>{ if(!priceSearch.trim())return prices; const q=priceSearch.toLowerCase(); return prices.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(q))); },[prices,priceSearch]);

  const leadTotalPages=Math.max(1,Math.ceil(filteredLeads.length/PAGE_SIZE));
  const priceTotalPages=Math.max(1,Math.ceil(filteredPrices.length/PAGE_SIZE));
  const paginatedLeads=filteredLeads.slice((leadPage-1)*PAGE_SIZE,leadPage*PAGE_SIZE);
  const paginatedPrices=filteredPrices.slice((pricePage-1)*PAGE_SIZE,pricePage*PAGE_SIZE);

  // ── Stats — berdasarkan filter aktif ──
  const stats=useMemo(()=>{
    const base=filteredLeads;
    const t=leads.length, tf=base.length;
    const cl=base.filter(r=>r["Status"]==="4 - Closed").length;
    const ht=base.filter(r=>r["Status"]==="3 - Hot").length;
    const fl=base.filter(r=>r["Status"]==="5 - Failed").length;
    const tv=base.reduce((s,r)=>s+(parseFloat(String(r["Value"]).replace(/\D/g,""))||0),0);
    const cr=t?((leads.filter(r=>r["Status"]==="4 - Closed").length/t)*100).toFixed(1):"0.0";
    const activeFilter=statFilter||filterStatus!=="all"||filterGroup!=="all"||Object.values(colFilters).some(Boolean)||leadSearch;
    return[
      {label:"Total Lead",value:t,sub:activeFilter?`${tf} tampil`:null,color:"#185fa5",bg:"#e6f1fb",filter:null},
      {label:"Closed",value:cl,color:"#3b6d11",bg:"#eaf3de",filter:"4 - Closed"},
      {label:"Hot 🔥",value:ht,color:"#a32d2d",bg:"#fcebeb",filter:"3 - Hot"},
      {label:"Failed",value:fl,color:"#791f1f",bg:"#fce8e8",filter:"5 - Failed"},
      {label:"Conversion",value:`${cr}%`,color:"#534ab7",bg:"#eeedfe",filter:null},
      {label:"Total Value",value:fmtValueShort(tv),sub:tv>0?`Rp ${tv.toLocaleString("id-ID")}`:null,color:"#854f0b",bg:"#faeeda",filter:null,wide:true},
    ];
  },[leads,filteredLeads,statFilter,filterStatus,filterGroup,colFilters,leadSearch]);

  if(dbError) return(<div style={{minHeight:"100vh",background:"#f4f6f9",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif",padding:20}}><div style={{background:"#fff",borderRadius:16,padding:"40px",maxWidth:520,textAlign:"center",boxShadow:"0 8px 32px rgba(0,0,0,0.08)"}}><div style={{fontSize:48,marginBottom:12}}>🔥</div><div style={{fontWeight:700,fontSize:18,marginBottom:12,color:"#d93025"}}>Firebase Belum Terhubung</div><div style={{fontSize:14,color:"#555",marginBottom:24,lineHeight:1.6}}>{dbError}</div><div style={{background:"#f8fafc",borderRadius:8,padding:"16px",textAlign:"left",fontSize:13,color:"#333",fontFamily:"monospace"}}><div style={{marginBottom:8,fontWeight:700,fontFamily:"sans-serif",color:"#666"}}>Langkah setup:</div>1. Buka <strong>Firebase Console</strong> → buat project baru<br/>2. Aktifkan <strong>Firestore Database</strong><br/>3. Buka Project Settings → Your apps → tambah Web app<br/>4. Copy config ke <strong>src/firebase.js</strong><br/>5. Di Firestore Rules: <code style={{background:"#eee",padding:"1px 4px"}}>allow read, write: if true;</code></div></div></div>);
  if(!role) return <LoginScreen onLogin={setRole}/>;
  if(showAdmin) return <AdminPanel master={master} setMasterField={setMasterField} trashedLeads={trashedLeads} onRestore={handleRestoreLead} onPermanentDelete={handlePermanentDelete} onBack={()=>setShowAdmin(false)}/>;

  const ConfirmDelete=({onConfirm,onCancel,label})=>(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:99999,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{background:"#fff",borderRadius:12,padding:"28px 32px",width:380,boxShadow:"0 16px 48px rgba(0,0,0,0.18)"}}><div style={{fontWeight:700,fontSize:16,marginBottom:6}}>Hapus {label}?</div><div style={{color:"#666",fontSize:14,marginBottom:6}}>Lead akan dipindahkan ke <strong>Sampah</strong>.</div><div style={{color:"#888",fontSize:13,marginBottom:24}}>Admin dapat restore atau hapus permanen di menu Master Data → Sampah.</div><div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button onClick={onCancel} style={{padding:"9px 18px",borderRadius:7,border:"1.5px solid #d0d5dd",background:"#f5f5f5",color:"#444",cursor:"pointer",fontSize:13}}>Batal</button><button onClick={onConfirm} style={{padding:"9px 18px",borderRadius:7,border:"none",background:"#d93025",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>🗑️ Pindah ke Sampah</button></div></div></div>);

  const activeFiltersCount=Object.values(colFilters).filter(Boolean).length+(statFilter?1:0)+(filterStatus!=="all"?1:0)+(filterGroup!=="all"?1:0);

  return (
    <div style={{minHeight:"100vh",background:"#f4f6f9",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif"}}>
      {toast&&<div style={{position:"fixed",top:16,right:16,zIndex:99999,background:toast.ok?"#eaf3de":"#fcebeb",color:toast.ok?"#2d6a11":"#a32d2d",border:`1px solid ${toast.ok?"#b3d97a":"#f0a0a0"}`,borderRadius:8,padding:"11px 18px",fontSize:13,fontWeight:500,boxShadow:"0 4px 16px rgba(0,0,0,0.10)",pointerEvents:"none"}}>{toast.ok?"✓ ":"✕ "}{toast.msg}</div>}
      {leadDeleteId&&<ConfirmDelete label="lead ini" onConfirm={()=>handleDeleteLead(leadDeleteId)} onCancel={()=>setLeadDeleteId(null)}/>}
      {priceDeleteId&&<ConfirmDelete label="data harga ini" onConfirm={()=>handleDeletePrice(priceDeleteId)} onCancel={()=>setPriceDeleteId(null)}/>}
      {leadModal&&(<Modal title={leadModal.type==="add"?"➕ Tambah Lead Baru":"✏️ Edit Lead"} onClose={()=>setLeadModal(null)} wide><LeadForm leads={allLeads} master={master} initial={leadModal.row||{}} isEdit={leadModal.type==="edit"} onSave={(form)=>handleSaveLead(form,leadModal.type==="edit",leadModal.row?._id)} onCancel={()=>setLeadModal(null)}/></Modal>)}
      {priceModal&&(<Modal title={priceModal.type==="add"?"➕ Tambah Data Harga":"✏️ Edit Data Harga"} onClose={()=>setPriceModal(null)} wide><PriceForm master={master} initial={priceModal.row||{}} isEdit={priceModal.type==="edit"} onSave={(form)=>handleSavePrice(form,priceModal.type==="edit",priceModal.row?._id)} onCancel={()=>setPriceModal(null)}/></Modal>)}

      {/* Header */}
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"0 24px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"13px 0 10px"}}>
          <span style={{fontSize:22}}>⚓</span>
          <div><div style={{fontWeight:800,fontSize:16,color:"#1a1a1a"}}>Lead Conversion Dashboard</div>
            <div style={{fontSize:12,color:"#888",display:"flex",alignItems:"center",gap:6}}>Marine Engine Sales Tracker<span style={{fontSize:11,background:"#eaf3de",color:"#2d6a11",padding:"1px 7px",borderRadius:10,fontWeight:600}}>{leadsLoading?"⏳ Memuat...":"🔥 Firebase Live"}</span></div>
          </div>
          <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
            <span style={{fontSize:11,background:role==="admin"?"#faeeda":"#e6f1fb",padding:"3px 9px",borderRadius:20,fontWeight:600,color:role==="admin"?"#854f0b":"#185fa5"}}>{role==="admin"?"🔐 Admin":"💼 Sales"}</span>
            {role==="admin"&&<button onClick={()=>setShowAdmin(true)} style={{padding:"7px 13px",borderRadius:8,border:"1.5px solid #185fa5",background:"#f0f7ff",color:"#185fa5",cursor:"pointer",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>⚙️ Master Data{trashedLeads.length>0&&<span style={{background:"#d93025",color:"#fff",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:700}}>{trashedLeads.length}</span>}</button>}
            <input ref={importRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{display:"none"}}/>
            <button onClick={()=>importRef.current?.click()} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 13px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#444",cursor:"pointer",fontSize:13,fontWeight:500}}>⬆️ Import Lead</button>
            <button onClick={handleExport} style={{display:"flex",alignItems:"center",gap:5,padding:"7px 13px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>⬇️ Export</button>
            <button onClick={()=>setRole(null)} style={{padding:"7px 11px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#888",cursor:"pointer",fontSize:12}}>Keluar</button>
          </div>
        </div>
        <div style={{display:"flex",gap:0}}>
          {[{id:"lead",icon:"📊",label:"Lead Tracker",count:leads.length},{id:"price",icon:"💰",label:"Product Price",count:prices.length}].map(t=>(<button key={t.id} onClick={()=>setActiveTab(t.id)} style={{padding:"9px 18px",fontSize:13,fontWeight:500,cursor:"pointer",border:"none",background:"transparent",borderBottom:activeTab===t.id?"2.5px solid #185fa5":"2.5px solid transparent",color:activeTab===t.id?"#185fa5":"#777",display:"flex",alignItems:"center",gap:7}}>{t.icon} {t.label}<span style={{background:activeTab===t.id?"#e6f1fb":"#f0f0f0",color:activeTab===t.id?"#185fa5":"#888",borderRadius:12,padding:"1px 8px",fontSize:11,fontWeight:600}}>{t.count}</span></button>))}
        </div>
      </div>

      <div style={{padding:"18px 24px"}}>
        {activeTab==="lead"&&(<>
          {/* Stats — clickable */}
          {leads.length>0&&(<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:14}}>
            {stats.map(s=>(
              <div key={s.label} onClick={()=>{if(s.filter){setStatFilter(statFilter===s.filter?null:s.filter);setLeadPage(1);}}} style={{background:s.bg,borderRadius:10,padding:"11px 14px",border:`1.5px solid ${statFilter===s.filter?s.color:s.color+"22"}`,cursor:s.filter?"pointer":"default",transition:"all 0.15s",boxShadow:statFilter===s.filter?"0 2px 8px rgba(0,0,0,0.12)":"none",position:"relative"}}>
                {statFilter===s.filter&&<span style={{position:"absolute",top:6,right:8,fontSize:10,color:s.color,fontWeight:700}}>✓ Filter aktif</span>}
                <div style={{fontSize:10,color:s.color,fontWeight:600,marginBottom:3,textTransform:"uppercase",letterSpacing:"0.04em"}}>{s.label}{s.filter?"  🔽":""}</div>
                <div style={{fontSize:19,fontWeight:700,color:s.color}}>{s.value}</div>
                {s.sub&&<div style={{fontSize:10,color:s.color,opacity:0.75,marginTop:2,fontFamily:"monospace"}}>{s.sub}</div>}
              </div>
            ))}
          </div>)}

          {/* Active filter bar */}
          {activeFiltersCount>0&&(<div style={{background:"#f0f7ff",border:"1px solid #b5d4f4",borderRadius:8,padding:"8px 14px",marginBottom:10,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"#185fa5",fontWeight:600}}>🔽 Filter aktif:</span>
            {statFilter&&<span style={{background:"#185fa5",color:"#fff",borderRadius:12,padding:"2px 10px",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>Status: {statFilter}<button onClick={()=>setStatFilter(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:13,padding:0,lineHeight:1,marginLeft:2}}>×</button></span>}
            {filterStatus!=="all"&&<span style={{background:"#185fa5",color:"#fff",borderRadius:12,padding:"2px 10px",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>Status: {filterStatus}<button onClick={()=>setFilterStatus("all")} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:13,padding:0,lineHeight:1,marginLeft:2}}>×</button></span>}
            {filterGroup!=="all"&&<span style={{background:"#185fa5",color:"#fff",borderRadius:12,padding:"2px 10px",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>Grup: {filterGroup}<button onClick={()=>setFilterGroup("all")} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:13,padding:0,lineHeight:1,marginLeft:2}}>×</button></span>}
            {Object.entries(colFilters).filter(([,v])=>v).map(([col,val])=>(<span key={col} style={{background:"#185fa5",color:"#fff",borderRadius:12,padding:"2px 10px",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:4}}>{col}: {val}<button onClick={()=>handleColFilter(col,"")} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:13,padding:0,lineHeight:1,marginLeft:2}}>×</button></span>))}
            <button onClick={()=>{setStatFilter(null);setFilterStatus("all");setFilterGroup("all");setColFilters({});setLeadSearch("");setLeadPage(1);}} style={{marginLeft:"auto",background:"none",border:"none",cursor:"pointer",color:"#a32d2d",fontSize:12,fontWeight:600}}>Hapus semua filter ×</button>
          </div>)}

          {/* Toolbar */}
          <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
            <input value={leadSearch} onChange={e=>{setLeadSearch(e.target.value);setLeadPage(1);}} placeholder="🔍  Cari lead..." style={{flex:"1 1 200px",maxWidth:260,padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none"}}/>
            {statusOptions.length>0&&<select value={filterStatus} onChange={e=>{setFilterStatus(e.target.value);setLeadPage(1);}} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#444",fontSize:13}}><option value="all">Semua Status</option>{statusOptions.map(s=><option key={s} value={s}>{s}</option>)}</select>}
            {groupOptions.length>0&&<select value={filterGroup} onChange={e=>{setFilterGroup(e.target.value);setLeadPage(1);}} style={{padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#444",fontSize:13}}><option value="all">Semua Grup</option>{groupOptions.map(g=><option key={g} value={g}>{g}</option>)}</select>}
            <span style={{fontSize:12,color:"#999"}}>{filteredLeads.length} lead</span>
            <button onClick={()=>setLeadModal({type:"add"})} style={{marginLeft:"auto",padding:"8px 16px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>＋ Lead Baru</button>
          </div>

          {leadsLoading?(<div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px",textAlign:"center",color:"#aaa",fontSize:14}}>⏳ Memuat data dari Firebase...</div>)
            :leads.length===0?(<div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px 24px",textAlign:"center"}}><div style={{fontSize:48,marginBottom:12}}>📂</div><div style={{fontWeight:700,fontSize:15,marginBottom:8}}>Belum ada lead</div><div style={{fontSize:13,color:"#888",marginBottom:24}}>Import Excel atau tambah lead baru.</div><div style={{display:"flex",gap:10,justifyContent:"center"}}><button onClick={()=>importRef.current?.click()} style={{padding:"9px 20px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>⬆️ Import Excel</button><button onClick={()=>setLeadModal({type:"add"})} style={{padding:"9px 20px",borderRadius:8,border:"1.5px solid #185fa5",background:"#fff",color:"#185fa5",cursor:"pointer",fontSize:13,fontWeight:600}}>＋ Tambah Lead</button></div></div>)
            :(<LeadTable rows={paginatedLeads} allRows={leads} page={leadPage} totalPages={leadTotalPages} filtered={filteredLeads} onEdit={row=>setLeadModal({type:"edit",row})} onDelete={id=>setLeadDeleteId(id)} role={role} onPageChange={setLeadPage} sortCol={sortCol} sortDir={sortDir} onSort={handleSort} colFilters={colFilters} onColFilter={handleColFilter}/>)
          }
        </>)}

        {activeTab==="price"&&(<>
          <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
            <input value={priceSearch} onChange={e=>{setPriceSearch(e.target.value);setPricePage(1);}} placeholder="🔍  Cari produk..." style={{flex:"1 1 200px",maxWidth:280,padding:"8px 12px",borderRadius:8,border:"1.5px solid #d0d5dd",background:"#fff",color:"#1a1a1a",fontSize:13,outline:"none"}}/>
            <span style={{fontSize:12,color:"#999"}}>{filteredPrices.length} produk</span>
            <button onClick={()=>setPriceModal({type:"add"})} style={{marginLeft:"auto",padding:"8px 16px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:700}}>＋ Tambah Harga</button>
          </div>
          {pricesLoading?(<div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px",textAlign:"center",color:"#aaa",fontSize:14}}>⏳ Memuat data dari Firebase...</div>)
            :prices.length===0?(<div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,padding:"60px 24px",textAlign:"center"}}><div style={{fontSize:48,marginBottom:12}}>💰</div><div style={{fontWeight:700,fontSize:15,marginBottom:8}}>Belum ada data harga</div><div style={{display:"flex",gap:10,justifyContent:"center",marginTop:20}}><button onClick={()=>importRef.current?.click()} style={{padding:"9px 20px",borderRadius:8,border:"none",background:"#185fa5",color:"#fff",cursor:"pointer",fontSize:13,fontWeight:600}}>⬆️ Import Excel</button><button onClick={()=>setPriceModal({type:"add"})} style={{padding:"9px 20px",borderRadius:8,border:"1.5px solid #185fa5",background:"#fff",color:"#185fa5",cursor:"pointer",fontSize:13,fontWeight:600}}>＋ Tambah Harga</button></div></div>)
            :(<PriceTable rows={paginatedPrices} page={pricePage} totalPages={priceTotalPages} filtered={filteredPrices} onEdit={row=>setPriceModal({type:"edit",row})} onDelete={id=>setPriceDeleteId(id)} role={role} onPageChange={setPricePage}/>)
          }
        </>)}
      </div>
      <style>{`button:disabled{opacity:0.4;cursor:not-allowed!important;} tr:last-child td{border-bottom:none!important;}`}</style>
    </div>
  );
}
