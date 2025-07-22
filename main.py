import streamlit as st
import io
from document_checker import DocumentChecker, DocumentStandardsParser
from docx import Document

def main(): 
    st.set_page_config(
        page_title="ตรวจสอบรูปแบบเอกสารรายงานโครงงาน",
        page_icon="📝",
        layout="wide"
    )
    
    st.title("🔍 ระบบตรวจสอบรูปแบบเอกสารรายงานโครงงาน")
    
    st.markdown("---")
    
    # อัปโหลดไฟล์ XML template (ไม่บังคับ)
    with st.expander("⚙️ การตั้งค่า XML Template (ไม่บังคับ)", expanded=False):
        st.markdown("**หากไม่อัปโหลด จะใช้ค่าเริ่มต้นจากระบบ**")
        xml_file = st.file_uploader(
            "เลือกไฟล์ XML Template",
            type="xml",
            help="อัปโหลดไฟล์ document_template.xml เพื่อกำหนดมาตรฐานการตรวจสอบ"
        )
        
        if xml_file is not None:
            st.success(f"✅ อัปโหลด XML template สำเร็จ: {xml_file.name}")
    
    # โหลด standards
    if xml_file is not None:
        # อ่านจากไฟล์ที่อัปโหลด
        xml_content = xml_file.read().decode('utf-8')
        standards = DocumentStandardsParser.parse_xml_content(xml_content)
        st.info("🔧 ใช้การตั้งค่าจากไฟล์ XML ที่อัปโหลด")
    else:
        # ใช้ค่าเริ่มต้น
        standards = DocumentStandardsParser._get_default_standards()
        st.info("🔧 ใช้การตั้งค่าเริ่มต้นของระบบ")

    # คำอธิบายวิธีการใช้งาน
    with st.expander("📋 วิธีการใช้งาน", expanded=False):
        st.markdown("""
        ### 🔄 ขั้นตอนการใช้งาน:
        1. **อัปโหลดไฟล์ XML Template** (ไม่บังคับ) เพื่อกำหนดมาตรฐานการตรวจสอบ
        2. **อัปโหลดไฟล์ Word (.docx)** ของรายงานโครงงาน
        3. **กดปุ่มตรวจสอบ** และรอผลการตรวจสอบ
        4. **ดาวน์โหลดไฟล์** ที่ตรวจสอบแล้ว (มีข้อผิดพลาดแสดงในเอกสาร)
        5. **เปิดไฟล์** Word และแก้ไขตามข้อแนะนำที่แสดงในเอกสาร
        
        ### ✨ คุณสมบัติของระบบ:
        - 🖼️ **ตรวจสอบการอ้างอิงรูปภาพ**: ตรวจสอบรูปแบบ "ภาพที่ X" และตำแหน่งการจัดแนว
        - 📊 **ตรวจสอบการอ้างอิงตาราง**: ตรวจสอบรูปแบบ "ตารางที่ X" และตำแหน่งการจัดแนว
        - 📏 **ตรวจสอบหน้าปกแบบละเอียด**: ตรวจสอบตามมาตรฐาน PDF ที่แนบมา
        - 🎯 **การตรวจสอบเฉพาะเจาะจง**: แยกประเภทข้อผิดพลาดได้ชัดเจนขึ้น
        - 📊 **สถิติที่ครบถ้วน**: แสดงข้อมูลการตรวจสอบและสถิติข้อผิดพลาดแบบละเอียด
        """)
    
    # อัปโหลดไฟล์เอกสาร
    st.markdown("### 📄 อัปโหลดเอกสารที่ต้องการตรวจสอบ")
    uploaded_file = st.file_uploader(
        "เลือกไฟล์เอกสาร Word (.docx)",
        type="docx",
        help="รองรับเฉพาะไฟล์ .docx เท่านั้น"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ อัปโหลดไฟล์สำเร็จ: {uploaded_file.name}")
        
        if st.button("🔍 เริ่มตรวจสอบเอกสาร", type="primary"):
            with st.spinner("🔄 กำลังตรวจสอบเอกสาร กรุณารอสักครู่..."):
                try:
                    # อ่านไฟล์ Word
                    doc = Document(uploaded_file)
                    
                    # สร้าง DocumentChecker
                    checker = DocumentChecker(standards)
                    
                    # ตรวจสอบเอกสาร
                    errors = checker.check_document_object(doc)
                    error_stats = checker.get_error_statistics()
                    
                    # สร้างไฟล์ที่มีการแสดงข้อผิดพลาด
                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    output_buffer.seek(0)
                    
                    # แสดงผลการตรวจสอบ
                    st.markdown("---")
                    st.header("📊 ผลการตรวจสอบ")
                    
                    # แสดงสถิติหลัก
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        total_paragraphs = len([p for p in doc.paragraphs if p.text.strip()])
                        st.metric("จำนวนย่อหน้าทั้งหมด", total_paragraphs)
                    with col2:
                        st.metric("จำนวนข้อผิดพลาด", len(errors))
                    with col3:
                        accuracy = ((total_paragraphs - len(errors)) / total_paragraphs * 100) if total_paragraphs else 100
                        st.metric("ความถูกต้อง", f"{accuracy:.1f}%")
                    with col4:
                        most_common_error = max(error_stats.items(), key=lambda x: x[1]) if error_stats else ("ไม่มี", 0)
                        st.metric("ข้อผิดพลาดหลัก", f"{most_common_error[1]} จุด")
                    
                    # แสดงสถิติประเภทข้อผิดพลาดแบบกราฟ
                    if error_stats:
                        st.subheader("📊 สถิติข้อผิดพลาดตามประเภท")
                        
                        # แบ่งข้อผิดพลาดตามหมวดหมู่
                        font_errors = {}
                        format_errors = {}
                        reference_errors = {}
                        structure_errors = {}
                        
                        for error_type, count in error_stats.items():
                            if any(keyword in error_type for keyword in ['ฟอนต์', 'ขนาด']):
                                font_errors[error_type] = count
                            elif any(keyword in error_type for keyword in ['ความหนา', 'จัดแนว', 'เยื้อง', 'ระยะห่าง']):
                                format_errors[error_type] = count
                            elif any(keyword in error_type for keyword in ['อ้างอิง', 'ตาราง', 'รูปภาพ']):
                                reference_errors[error_type] = count
                            else:
                                structure_errors[error_type] = count
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if font_errors:
                                st.write("**🔤 ข้อผิดพลาดเกี่ยวกับฟอนต์:**")
                                for error_type, count in sorted(font_errors.items(), key=lambda x: x[1], reverse=True):
                                    progress = count / max(error_stats.values()) if error_stats.values() else 0
                                    st.progress(progress, text=f"{error_type}: {count} จุด")
                            
                            if reference_errors:
                                st.write("**📊 ข้อผิดพลาดเกี่ยวกับการอ้างอิง:**")
                                for error_type, count in sorted(reference_errors.items(), key=lambda x: x[1], reverse=True):
                                    progress = count / max(error_stats.values()) if error_stats.values() else 0
                                    st.progress(progress, text=f"{error_type}: {count} จุด")
                        
                        with col2:
                            if format_errors:
                                st.write("**📐 ข้อผิดพลาดเกี่ยวกับการจัดรูปแบบ:**")
                                for error_type, count in sorted(format_errors.items(), key=lambda x: x[1], reverse=True):
                                    progress = count / max(error_stats.values()) if error_stats.values() else 0
                                    st.progress(progress, text=f"{error_type}: {count} จุด")
                            
                            if structure_errors:
                                st.write("**🏗️ ข้อผิดพลาดเกี่ยวกับโครงสร้าง:**")
                                for error_type, count in sorted(structure_errors.items(), key=lambda x: x[1], reverse=True):
                                    progress = count / max(error_stats.values()) if error_stats.values() else 0
                                    st.progress(progress, text=f"{error_type}: {count} จุด")
                        
                        # สรุปข้อผิดพลาดทั้งหมด
                        st.subheader("📝 รายละเอียดข้อผิดพลาดทั้งหมด")
                        for error_type, count in sorted(error_stats.items(), key=lambda x: x[1], reverse=True):
                            st.write(f"• **{error_type}**: {count} จุด")
                    
                    # ปุ่มดาวน์โหลดไฟล์ที่ตรวจสอบแล้ว
                    if errors:
                        st.markdown("---")
                        st.subheader("📥 ดาวน์โหลดไฟล์ที่ตรวจสอบแล้ว")
                        
                        col1, col2 = st.columns([2, 1])
                        with col1:
                            st.download_button(
                                label="📄 ดาวน์โหลดเอกสารที่มีข้อผิดพลาดแสดงในเอกสาร",
                                data=output_buffer.getvalue(),
                                file_name=f"checked_{uploaded_file.name}",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary"
                            )
                        with col2:
                            st.metric("ไฟล์ขนาด", f"{len(output_buffer.getvalue()) / 1024:.1f} KB")
                        
                        st.info("💡 **คำแนะนำ**: เปิดไฟล์ที่ดาวน์โหลด จะเห็นข้อความที่มีปัญหาถูก highlight สีเหลือง และมีข้อความสีแดงขนาด 12pt บอกวิธีแก้ไข")
                        
                        # แสดงตัวอย่างข้อผิดพลาดที่พบ
                        if len(errors) > 0:
                            st.subheader("🔍 ตัวอย่างข้อผิดพลาดที่พบ")
                            sample_errors = errors[:5]  # แสดง 5 รายการแรก
                            for i, error in enumerate(sample_errors, 1):
                                with st.expander(f"ข้อผิดพลาดที่ {i}: {error['type']}"):
                                    st.write(f"**ย่อหน้าที่**: {error['paragraph']}")
                                    st.write(f"**ประเภท**: {error['type']}")
                                    st.write(f"**รายละเอียด**: {error['description']}")
                            
                            if len(errors) > 5:
                                st.info(f"และอีก {len(errors) - 5} ข้อผิดพลาดอื่นๆ ในไฟล์ที่ดาวน์โหลด")
                        
                    else:
                        st.success("🎉 **ยินดีด้วย!** เอกสารของคุณผ่านการตรวจสอบทุกข้อแล้ว")
                        
                        # แสดงข้อมูลสถิติเมื่อผ่านการตรวจสอบ
                        st.balloons()
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("✅ การตรวจสอบเสร็จสิ้น", "100%")
                        with col2:
                            st.metric("📄 ย่อหน้าที่ถูกต้อง", total_paragraphs)
                        with col3:
                            st.metric("⭐ คะแนน", "A+")
                        
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาดในการตรวจสอบ: {str(e)}")
                    
                    with st.expander("🔧 คำแนะนำการแก้ไขปัญหา", expanded=True):
                        st.write("**ขั้นตอนการแก้ไข:**")
                        st.write("1. ตรวจสอบว่าไฟล์เป็น .docx หรือไม่")
                        st.write("2. ลองปิดไฟล์ใน Microsoft Word ก่อนอัปโหลด")
                        st.write("3. ตรวจสอบว่าไฟล์ไม่เสียหาย")
                        st.write("4. หากใช้ XML template ตรวจสอบว่าไฟล์ XML ถูกต้อง")
                        st.write("5. ลองใช้ไฟล์เอกสารตัวอย่างที่แน่ใจว่าทำงานได้")
                        
                        st.write("**ข้อมูลสำหรับการแก้ไข:**")
                        st.code(f"ข้อผิดพลาด: {str(e)}")
    
    # คำแนะนำเพิ่มเติม
    st.markdown("---")
    with st.expander("💡 เทคนิคการจัดรูปแบบเอกสารรายงานโครงงาน"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **🎯 วิธีปรับแต่งฟอนต์ทั้งเอกสาร:**
            1. กด Ctrl+A เพื่อเลือกทั้งเอกสาร
            2. ไปที่ Home > Font > เลือก TH Sarabun New
            3. ปรับขนาดฟอนต์ตามประเภทเนื้อหา
            
            **📐 การจัดการเยื้อง:**
            - เนื้อหา: กด Tab หรือใช้ Paragraph > Special > First line
            - หัวข้อ: ไม่ต้องเยื้อง ให้ชิดซ้าย
            
            **⚖️ ขนาดฟอนต์ตามมาตรฐาน:**
            - ชื่อบท (บทที่, สารบัญ): 18pt, Bold
            - หัวข้อย่อย (1.1, 2.1): 16pt, Bold
            - เนื้อหาทั่วไป: 14pt, Regular
            - ข้อมูลในหน้าปก: 14pt
            - ชื่อเรื่องในหน้าปก: 18pt
            """)
        
        with col2:
            st.markdown("""
            **📊 การอ้างอิงตารางและรูปภาพ:**
            - ตารางที่ 1, ตารางที่ 2.1 (ด้านบน-ชิดซ้าย)
            - ภาพที่ 1, ภาพที่ 2.1 (ด้านล่าง-กึ่งกลาง)
            - ใช้ขนาดฟอนต์ 14pt แบบปกติ
            
            **📄 การจัดรูปแบบหน้าปก:**
            - หน้าปกประกอบด้วย 2 หน้า (ปกนอกและปกใน)
            - ข้อมูลต้องจัดกึ่งกลาง
            - ไม่ต้องเยื้องย่อหน้า
            - ใช้รหัสนักศึกษารูปแบบ 9 หลัก-1 หลัก
            
            **🔍 เคล็ดลับการตรวจสอบ:**
            - ใช้ Find & Replace เพื่อแก้ไขฟอนต์ทั้งเอกสาร
            - ตรวจสอบทีละส่วน (ปก, สารบัญ, เนื้อหา)
            - Save ไฟล์เป็น .docx ให้แน่ใจ
            """)
    
    # คำแนะนำการใช้งานเอกสารที่ตรวจสอบแล้ว
    with st.expander("📖 วิธีการใช้เอกสารที่ตรวจสอบแล้ว"):
        st.markdown("""
        **📝 เมื่อเปิดเอกสารที่ดาวน์โหลด จะเห็น:**
        
        1. **จุดผิดพลาดในเนื้อหา**
            - ข้อความที่มีปัญหาจะมี **highlight สีเหลือง**
            - ข้อความสีแดงขนาด 12pt บอกข้อผิดพลาดจะปรากฏที่ท้ายย่อหน้า
            - รูปแบบ: `[❌ ข้อผิดพลาด: รายละเอียดข้อผิดพลาด]`
            - **แต่ละย่อหน้าจะแสดงเพียงครั้งเดียว** ไม่ซ้ำ
        
        **🔧 ขั้นตอนการแก้ไข:**
        1. ใช้ Ctrl+F ค้นหาข้อความ "❌ ข้อผิดพลาด" เพื่อหาจุดผิดพลาด
        2. แก้ไขข้อผิดพลาดตามคำแนะนำ
        3. ลบข้อความแสดงข้อผิดพลาด (ข้อความสีแดง) ออก
        4. ลบ highlight สีเหลือง (เลือกข้อความ > Home > Text Highlight Color > No Color)
        
        **⚠️ ข้อควรระวัง:**
        - อย่าลืมลบข้อความแสดงข้อผิดพลาดออกหลังแก้ไขเสร็จ
        - ตรวจสอบให้แน่ใจว่าแก้ไขถูกต้องก่อนลบข้อความแสดงข้อผิดพลาด
        - สำรองเอกสารต้นฉบับไว้ก่อนแก้ไข
        """)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>Document Checker for Project Reports v6.2 | Cleaned Version | Powered by Streamlit</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()