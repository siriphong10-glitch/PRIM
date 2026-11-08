# PRIM — Project & Resource Integrated Management

เว็บแอปบริหารแผนงานโครงการและกำลังคน (single-file HTML) สำหรับ Performax Building Service

- `index.html` — ตัวโปรแกรม (rev 08.11.26) เปิดผ่าน GitHub Pages
- `backend/Code.gs` — ฐานข้อมูลกลาง + ระบบเข้าสู่ระบบบน Google Sheets (Apps Script Web App)
- `legacy-SiriPRIM-rev30.04.26.html` — โปรแกรมรุ่นเดิม (เก็บไว้อ้างอิง)

## ติดตั้งฐานข้อมูลกลาง (ครั้งเดียว)
1. เปิด script.google.com/create ด้วยบัญชี Google ของบริษัท → วางโค้ดจาก `backend/Code.gs` → บันทึก (สคริปต์จะสร้าง Google Sheet "PRIM Database" ให้เอง)
2. การทำให้ใช้งานได้ → รายการใหม่ → เว็บแอป → ดำเนินการในฐานะ "ฉัน" → ผู้มีสิทธิ์เข้าถึง "ทุกคน" → คัดลอก URL `/exec`
3. เปิดเว็บ PRIM → "เชื่อมต่อฐานข้อมูลกลาง" → วาง URL → ผู้ที่เข้าสู่ระบบคนแรกเป็นผู้ดูแลระบบ
4. เมนูชื่อผู้ใช้ → ตั้งค่าการเชื่อมต่อ → คัดลอก "ลิงก์สำหรับทีม" ส่งให้ทีม
