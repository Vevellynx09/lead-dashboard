# 🚀 Panduan Deploy: GitHub + Firebase + Vercel

## Hasil akhir: App kamu online permanen, database Firebase, auto-update setiap push ke GitHub

---

## BAGIAN 1 — Setup Firebase (Database)

### Langkah 1: Buat Firebase Project
1. Buka https://console.firebase.google.com
2. Klik **"Add project"** → isi nama project (contoh: `lead-dashboard`)
3. Klik **Continue** → matikan Google Analytics (opsional) → **Create project**

### Langkah 2: Aktifkan Firestore
1. Di sidebar kiri → klik **"Firestore Database"**
2. Klik **"Create database"**
3. Pilih **"Start in test mode"** → klik **Next**
4. Pilih region terdekat (contoh: `asia-southeast1`) → klik **Enable**

### Langkah 3: Ambil Konfigurasi Firebase
1. Klik ikon ⚙️ (Settings) → **"Project settings"**
2. Scroll ke bawah → bagian **"Your apps"** → klik ikon **Web (</>)**
3. Isi nama app (bebas) → klik **"Register app"**
4. **COPY semua isi `firebaseConfig`** — kamu akan butuhkan ini di langkah berikutnya

Contoh yang akan kamu copy:
```
apiKey: "AIzaSy...",
authDomain: "lead-dashboard-xxxx.firebaseapp.com",
projectId: "lead-dashboard-xxxx",
storageBucket: "lead-dashboard-xxxx.appspot.com",
messagingSenderId: "123456789",
appId: "1:123456789:web:abcdef"
```

### Langkah 4: Atur Firestore Rules (izin akses)
1. Di Firestore → klik tab **"Rules"**
2. Hapus semua teks yang ada, ganti dengan ini:
```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if true;
    }
  }
}
```
3. Klik **"Publish"**

---

## BAGIAN 2 — Upload Kode ke GitHub

### Langkah 5: Buat GitHub Account (kalau belum punya)
- Daftar di https://github.com

### Langkah 6: Install Git (kalau belum)
- Download di https://git-scm.com/downloads → install → restart PC

### Langkah 7: Buat Repository GitHub
1. Buka https://github.com/new
2. Repository name: `lead-dashboard`
3. Pilih **Private** (agar kode tidak publik)
4. Klik **"Create repository"**

### Langkah 8: Upload Project ke GitHub
Buka **Terminal / Command Prompt**, masuk ke folder project:

```bash
cd lead-dashboard
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/USERNAME_KAMU/lead-dashboard.git
git push -u origin main
```

> Ganti `USERNAME_KAMU` dengan username GitHub kamu

---

## BAGIAN 3 — Deploy ke Vercel (Hosting Gratis)

### Langkah 9: Buat Vercel Account
1. Buka https://vercel.com
2. Klik **"Sign up"** → pilih **"Continue with GitHub"**
3. Izinkan akses → account otomatis terhubung ke GitHub

### Langkah 10: Import Project
1. Di Vercel dashboard → klik **"Add New"** → **"Project"**
2. Cari repo `lead-dashboard` → klik **"Import"**
3. Framework preset: **Vite** (biasanya otomatis terdeteksi)
4. **JANGAN klik Deploy dulu** — lanjut ke langkah 11

### Langkah 11: Tambahkan Firebase Config sebagai Environment Variables
Masih di halaman yang sama, scroll ke **"Environment Variables"**:

Tambahkan satu per satu (nama = Value dari firebase config kamu):

| NAME | VALUE |
|------|-------|
| `VITE_FIREBASE_API_KEY` | `AIzaSy...` (isi dari config kamu) |
| `VITE_FIREBASE_AUTH_DOMAIN` | `project-id.firebaseapp.com` |
| `VITE_FIREBASE_PROJECT_ID` | `project-id` |
| `VITE_FIREBASE_STORAGE_BUCKET` | `project-id.appspot.com` |
| `VITE_FIREBASE_MESSAGING_SENDER_ID` | `123456789` |
| `VITE_FIREBASE_APP_ID` | `1:123...` |

### Langkah 12: Deploy!
1. Klik **"Deploy"**
2. Tunggu 1-2 menit
3. Vercel akan kasih URL seperti: **`https://lead-dashboard-xxxx.vercel.app`** ✅

---

## BAGIAN 4 — Setelah Deploy

### Cara Update App di Masa Depan
Setiap kali kamu edit file dan push ke GitHub, Vercel **otomatis deploy ulang**:
```bash
git add .
git commit -m "update fitur baru"
git push
```
Selesai! App online dalam 1-2 menit.

### Cara Lihat & Manage Data
- **Firebase Console** → Firestore Database → bisa lihat, edit, hapus data langsung
- Collections yang dipakai: `leads`, `prices`, `config`

### Cara Tambah Admin Baru
App ini pakai role sederhana (Admin / Sales) — tidak perlu Firebase Auth.
Kalau butuh login dengan password, kabari untuk ditambahkan fitur auth.

---

## RINGKASAN LINK PENTING

| Keperluan | Link |
|-----------|------|
| Firebase Console | https://console.firebase.google.com |
| GitHub | https://github.com |
| Vercel Dashboard | https://vercel.com/dashboard |
| App kamu (setelah deploy) | Lihat di Vercel dashboard |

---

## ❓ Troubleshooting

**App error "Firebase Belum Terhubung"**
→ Cek Environment Variables di Vercel sudah benar semua

**Data tidak muncul**
→ Cek Firestore Rules sudah dipublish

**Build error di Vercel**
→ Pastikan framework preset = Vite, bukan Create React App
