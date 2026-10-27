# TRC Rank สินค้า (AV 6M) – Streamlit Web App

มี 2 เมนู
1. **รวมข้อมูล** – อัปโหลด PPS, PLR, New Product, BarBom, Data Rank เดือนก่อน → ได้ Data Rank สมบูรณ์ (เติมยี่ห้อ/รุ่นอัตโนมัติ) (`scripts/merge_data_rank.py`)
2. **จัด Rank** – อัปโหลด Data Rank (จากข้อ 1), ไฟล์ Rank เทมเพลตของเดือน, CSR → คำนวณและจัด Rank A/B/C/D (`scripts/build_rank.py`)

## รันในเครื่อง
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Community Cloud)
1. สร้าง GitHub repo (แนะนำ **Private**) แล้ว push โฟลเดอร์นี้ขึ้นไป
2. เข้า https://share.streamlit.io → New app → เลือก repo/branch → Main file path = `app.py` → Deploy
3. (ไม่บังคับ) ตั้งรหัสผ่าน: App → Settings → Secrets ใส่ `APP_PASSWORD = "รหัสของคุณ"`

## ข้อควรระวัง
- ห้าม commit ไฟล์ข้อมูลจริง (.gitignore กันไว้แล้ว)
- ไฟล์ Rank/Master ใหญ่ (~70MB) ใช้ RAM ราว 1.5GB และใช้เวลาหลายนาที – Community Cloud (RAM ~1GB) อาจไม่พอ
  ถ้าเต็ม ให้ใช้เครื่องตัวเอง / VM / Streamlit แบบจ่ายเงิน / Hugging Face Spaces
- ไฟล์ยอดขายแคช `sales.json.gz` ดาวน์โหลดเก็บไว้ แล้วอัปโหลดเดือนหน้าเพื่อข้ามการอ่านไฟล์ใหญ่
- ไม่ได้ recalc ด้วย LibreOffice (ใช้ `--no-recalc`) ไฟล์จะคำนวณสูตรเมื่อเปิดใน Excel
