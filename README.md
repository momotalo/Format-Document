# 📝 ระบบตรวจสอบรูปแบบเอกสาร (Document Checker)

ระบบตรวจสอบรูปแบบเอกสาร Word อัตโนมัติ โดยใช้ XML template เป็นมาตรฐานในการตรวจสอบ

## 🎯 ฟีเจอร์หลัก

### การตรวจสอบรูปแบบเอกสาร

- **ฟอนต์**: ตรวจสอบ Theme Font (TH Sarabun New)
- **ขนาดฟอนต์**:
  - หัวข้อหลัก: 18pt
  - หัวข้อย่อย: 16pt
  - เนื้อหา: 14pt
- **การจัดรูปแบบ**: Bold สำหรับหัวข้อ, Regular สำหรับเนื้อหา
- **การเยื้อง**: ตรวจสอบ Indentation ของย่อหน้า
- **การอ้างอิง**: ตำแหน่งการอ้างอิงรูปภาพและตาราง

### ผลลัพธ์

- Highlight จุดผิดพลาดในเอกสาร
- รายงานข้อผิดพลาดแบบละเอียด
- ดาวน์โหลดเอกสารที่ปรับปรุงแล้ว

## 🚀 การติดตั้งและใช้งาน

### ความต้องการของระบบ

- Python 3.8 หรือสูงกว่า
- pip (Python package manager)

### ขั้นตอนการติดตั้ง

1. **Clone หรือดาวน์โหลดโปรเจค**

```bash
git clone [https://github.com/momotalo/Format-Document.git]
cd Format Document
```

2. **สร้าง Virtual Environment (แนะนำ)**

```bash
python -m venv venv
source venv/bin/activate  # สำหรับ Linux/Mac
# หรือ
venv\Scripts\activate     # สำหรับ Windows

#หากต้องการออกจากโหลด
deactivate 
```

3. **ติดตั้ง Dependencies**

```bash
pip install -r requirements.txt
```

4. **รันแอพพลิเคชัน**

```bash
streamlit run main.py
```

5. **เปิดเบราว์เซอร์**
   - แอพจะเปิดอัตโนมัติที่ `http://localhost:8501`
   - หากไม่เปิดอัตโนมัติ ให้คลิกลิงก์ที่แสดงใน Terminal

## 📋 วิธีการใช้งาน

### 1. อัปโหลดไฟล์

- คลิก "เลือกไฟล์ Word (.docx)"
- เลือกไฟล์เอกสาร Word ที่ต้องการตรวจสอบ

### 2. ตรวจสอบเอกสาร

- คลิกปุ่ม "🔍 เริ่มตรวจสอบเอกสาร"
- รอให้ระบบประมวลผล

### 3. ดูผลการตรวจสอบ

- ระบบจะแสดงรายการข้อผิดพลาด (หากมี)
- ดูรายละเอียดในแต่ละจุดที่ผิดพลาด

### 4. ดาวน์โหลดผลลัพธ์

- คลิกปุ่ม "📥 ดาวน์โหลดเอกสารที่ตรวจสอบแล้ว"
- เอกสารจะมี highlight สีเหลืองที่จุดผิดพลาด

## 📁 โครงสร้างไฟล์

```
document-checker/
├── main.py                 # ไฟล์หลักของแอพพลิเคชัน
├── document_template.xml   # XML template สำหรับมาตรฐานการตรวจสอบ
├── requirements.txt        # รายการ Python packages ที่ต้องใช้
└── README.md              # คำแนะนำการใช้งาน
```

## 🔧 การแก้ไขปัญหา

### ปัญหาที่พบบ่อย

1. **ไม่สามารถติดตั้ง python-docx**

   ```bash
   pip install --upgrade pip
   pip install python-docx
   ```

2. **Streamlit ไม่ทำงาน**

   ```bash
   pip install --upgrade streamlit
   streamlit run main.py
   ```

3. **ไฟล์ XML ไม่ถูกต้อง**
   - ตรวจสอบ syntax XML ให้ถูกต้อง
   - ใช้ XML validator online

## 📊 ตัวอย่างผลการตรวจสอบ

### ข้อผิดพลาดที่พบบ่อย

- ฟอนต์ไม่ใช่ TH Sarabun New
- ขนาดฟอนต์ไม่ตรงตามมาตรฐาน
- หัวข้อไม่เป็น Bold
- ไม่มีการเยื้องในย่อหน้า

### สถิติที่แสดง

- จำนวนย่อหน้าทั้งหมด
- จำนวนข้อผิดพลาด
- เปอร์เซ็นต์ความถูกต้อง

---

**Document Checker v1.0** | สร้างด้วย Python & Streamlit
