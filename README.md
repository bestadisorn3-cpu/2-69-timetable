# ตารางเรียน · ตารางสอน แผนกวิชาช่างไฟฟ้า วิทยาลัยเทคนิคปราจีนบุรี

เว็บแสดงตารางเรียน (เลือก ปวช./ปวส. → ชั้นปี → ห้อง → กลุ่ม เฉพาะ ปวช.3) และตารางสอนรายครู
ภาคเรียนที่ 2/2569 — เว็บนิ่ง (static) ไม่ต้องมีเซิร์ฟเวอร์ ใช้ GitHub Pages ได้ทันที

## ขึ้น GitHub Pages
1. สร้าง repository ใหม่ (Public) เช่น `timetable`
2. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ (index.html, data.js, logo.png, favicon.png, apple-touch-icon.png, .nojekyll, assets/, tools/) ไว้ที่ root ของ repo
3. Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)` → Save
4. รอ 1–2 นาที เว็บจะอยู่ที่ `https://<ชื่อผู้ใช้>.github.io/timetable/`

ลิงก์ตรงรายห้อง (ทำ QR ติดหน้าห้องได้): `.../#/class/v3-1-g1` (ปวช.3/1 กลุ่ม 1), `.../#/class/s1-4` (ปวส.1/4)
ลิงก์ตารางสอนครู: `.../#/teacher/อ.อดิศร`

## อัปเดตตารางเทอมใหม่
ใช้ไฟล์ Excel รูปแบบเดิม (ชีตละห้อง 4 แถวต่อวัน) แล้วรัน
```
pip install openpyxl
python tools/build_data.py "ตารางเรียน.xlsx" data.js
```
จากนั้นเติมรายชื่อครูของแผนก (ตารางครูแสดงเฉพาะชื่อในไฟล์นี้) และคาบที่สอนแผนกอื่น
```
python tools/build_teachers.py "ตารางสอนรายบุคคล.xlsx" data.js
```
สคริปต์จะสร้าง `data.js` ใหม่ และพิมพ์รายการที่ควรตรวจ (ชั่วโมงไม่ตรง ท+ป, ห้องว่าง ฯลฯ) จากนั้นอัปโหลด `data.js` ทับไฟล์เดิมใน repo
