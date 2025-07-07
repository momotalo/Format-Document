# 1. เลือก Gase Image
FROM python:3.11-slim

# 2. ตั่งค่า Working directory ภายใน Image
WORKDIR /main

# 3. คัดลอกไฟล์ที่จำเป็นไปยัง Image
COPY requirements.txt .

# 4. ติดตั้ง Dependencies
RUN pip install -r requirements.txt

# 5. คัดลอกโค้ดไปยัง Image
COPY . .

# 6. กำหนดคำสั่งเริ่มต้นเมื่อรัน Container
CMD ["streamlit","run", "main.py"]