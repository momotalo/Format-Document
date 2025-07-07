from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import xml.etree.ElementTree as ET

class DocumentStandardsParser:
    """คลาสสำหรับอ่าน XML standards"""
    
    @staticmethod
    def parse_xml_file(xml_file_path):
        """อ่านไฟล์ XML และแปลงเป็น dictionary"""
        try:
            tree = ET.parse(xml_file_path)
            root = tree.getroot()
            return DocumentStandardsParser._parse_xml_root(root)
        except Exception as e:
            print(f"Error reading XML file: {e}")
            return DocumentStandardsParser._get_default_standards()
    
    @staticmethod
    def parse_xml_content(xml_content):
        """อ่าน XML content และแปลงเป็น dictionary"""
        try:
            root = ET.fromstring(xml_content)
            return DocumentStandardsParser._parse_xml_root(root)
        except Exception as e:
            print(f"Error parsing XML content: {e}")
            return DocumentStandardsParser._get_default_standards()
    
    @staticmethod
    def _parse_xml_root(root):
        """แปลง XML root element เป็น dictionary"""
        return {
            'theme_font': root.find('.//theme_font').text,
        'main_heading_size': int(root.find('.//main_heading').text),
        'content_size': int(root.find('.//content').text),
        'sub_heading_size': int(root.find('.//sub_heading').text),
        'cover_title_size': int(root.find('.//cover_title').text),
        'cover_info_size': int(root.find('.//cover_info').text),
        
        # Formatting rules
        'main_heading_bold': root.find('.//main_heading_bold').text.lower() == 'true',
        'sub_heading_bold': root.find('.//sub_heading_bold').text.lower() == 'true',
        'content_bold': root.find('.//content_bold').text.lower() == 'true',
        'cover_title_bold': root.find('.//cover_title_bold').text.lower() == 'true',
        
        # เพิ่มรองรับ major_heading (หัวข้อใหญ่) ที่ใช้ขนาด sub_heading
        'major_heading_size': int(root.find('.//sub_heading').text),  # ใช้ค่าเดียวกับ sub_heading
        'major_heading_bold': root.find('.//sub_heading_bold').text.lower() == 'true',
        
        # Indentation rules
        'paragraph_indent': root.find('.//paragraph_indent').text.lower() == 'true',
        'sub_heading_indent': root.find('.//sub_heading_indent').text.lower() == 'true',
        
        # Alignment rules
        'cover_center': root.find('.//cover_center').text.lower() == 'true',
        'main_heading_left': root.find('.//main_heading_left').text.lower() == 'true',
        'content_justify': root.find('.//content_justify').text.lower() == 'true',
        
        # References
        'picture_position': root.find('.//picture_position').text if root.find('.//picture_position') is not None else 'below_center',
        'table_position': root.find('.//table_position').text if root.find('.//table_position') is not None else 'above_left',
        
        # Section lists
        'main_headings': [item.text for item in root.findall('.//main_headings/item')],
        'cover_sections': [item.text for item in root.findall('.//cover_sections/item')],
        'no_indent_sections': [item.text for item in root.findall('.//no_indent_sections/item')],
        'center_sections': [item.text for item in root.findall('.//center_sections/item')],
        'sub_heading_patterns': [item.text for item in root.findall('.//sub_heading_patterns/pattern')],
        'special_spacing_sections': [item.text for item in root.findall('.//special_spacing_sections/item')] if root.find('.//special_spacing_sections') is not None else [],
        
        # Exceptions
        'skip_font_check': [item.text for item in root.findall('.//skip_font_check/item')],
        'skip_bold_check': [item.text for item in root.findall('.//skip_bold_check/item')]
        }
    
    @staticmethod
    def _get_default_standards():
        """ค่า default หากไม่สามารถอ่าน XML ได้ - อัปเดตให้ตรงกับ XML template"""
        return {
            'theme_font': 'TH Sarabun New',
            'main_heading_size': 18,  # ชื่อบท
            'major_heading_size': 16,  # หัวข้อใหญ่ (1. 2. 3.)
            'content_size': 14,
            'sub_heading_size': 16,   # เก็บไว้เพื่อความเข้ากันได้
            'cover_title_size': 18,
            'cover_info_size': 14,
            'main_heading_bold': True,
            'major_heading_bold': True,  # หัวข้อใหญ่เป็น Bold
            'sub_heading_bold': True,  # เก็บไว้เพื่อความเข้ากันได้
            'content_bold': False,
            'cover_title_bold': True,
            'paragraph_indent': True,
            'sub_heading_indent': False,
            'cover_center': True,
            'main_heading_left': True,
            'content_justify': False,
            'picture_position': 'below_center',
            'table_position': 'above_left',
            'main_headings': ['บทคัดย่อ', 'Abstract', 'กิตติกรรมประกาศ', 'สารบัญ', 'บทที่ 1', 'บทที่ 2', 'บทที่ 3', 'บทที่ 4', 'บทที่ 5', 'เอกสารอ้างอิง', 'ภาคผนวก', 'ประวัติผู้เขียน'],
            'cover_sections': ['เอกสารโครงงานฉบับสมบูรณ์', 'คู่มือการใช้งานระบบ', 'CS/IT', 'โดย', 'อาจารย์ที่ปรึกษา'],
            'no_indent_sections': ['ปก', 'สารบัญ', 'สารบัญตาราง', 'สารบัญภาพ', 'บทคัดย่อ', 'Abstract', 'กิตติกรรมประกาศ'],
            'center_sections': ['ปก', 'ภาคผนวก'],
            'sub_heading_patterns': [r'^\d+\.\d+\s', r'^\d+\.\d+\.\d+\s', r'^\(\d+\)\s'],
            'special_spacing_sections': ['บทคัดย่อ', 'Abstract'],
            'skip_font_check': ['หมายเลขหน้า', 'Header', 'Footer'],
            'skip_bold_check': ['Abstract', 'บทคัดย่อ', 'Acknowledgement']
        }

class DocumentChecker:
    """คลาสหลักสำหรับตรวจสอบเอกสาร"""
    
    def __init__(self, standards):
        self.standards = standards
        self.errors = []
        self.tables = []
        self.pictures = []
    
    def check_document(self, doc_path):
        """ตรวจสอบเอกสาร Word"""
        doc = Document(doc_path)
        self.errors = []
        self.tables = []
        self.pictures = []

        # ตรวจหาตารางและรูปภาพในเอกสาร
        self._extract_tables_and_pictures(doc)

        # ตรวจสอบหน้าปกตามมาตรฐาน PDF - แก้ไขให้ตรวจสอบเฉพาะขนาดฟอนต์และความหนา
        self._check_cover_page_detailed(doc)

        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():  # ข้ามย่อหน้าว่าง
                # ข้ามหน้าปกในการตรวจสอบปกติ
                if not self._is_cover_page_section(paragraph.text, i):
                    self._check_paragraph(paragraph, i + 1, i)

        # ตรวจสอบการอ้างอิงตารางและรูปภาพ - แก้ไขให้ตรวจสอบจาก caption style
        self._check_references(doc)

        return self.errors, doc

    def _extract_tables_and_pictures(self, doc):
        """ดึงข้อมูลตารางและรูปภาพจากเอกสาร"""
        table_count = 0
        for table in doc.tables:
            table_count += 1
            self.tables.append({
                'number': table_count,
                'position': None
            })

    def _is_main_heading(self, text, paragraph=None):
        """ตรวจสอบว่าเป็นชื่อบท (18pt Bold) หรือไม่ - แก้ไขให้แยกแยะจากหัวข้อใหญ่ได้ชัดเจน"""
        text_clean = text.strip()
        
        # ตรวจสอบ style ก่อน - Header 1 = ชื่อบท, Header 2 = หัวข้อใหญ่
        if paragraph:
            if self._has_header1_style(paragraph):
                return True
            elif self._has_header2_style(paragraph):
                return False  # Header 2 คือหัวข้อใหญ่ ไม่ใช่ชื่อบท

        def _is_acknowledgement_content(self, text):
            """ตรวจสอบว่าเป็นเนื้อหาในกิตติกรรมประกาศหรือไม่"""
            acknowledgement_content_patterns = [
                'ขอขอบพระคุณ', 'ขอขอบคุณ', 'ขอกราบขอบพระคุณ', 
                'คุณพ่อ', 'คุณแม่', 'ที่ให้คำปรึกษา', 'ที่ให้การสนับสนุน',
                'ครอบครัว', 'เพื่อน', 'อาจารย์', 'ที่ปรึกษา', 'คณาจารย์',
                'แนวทางในการ', 'ให้คำแนะนำ', 'และการพัฒนา', 'ในการเรียน',
                'ผู้เขียน', 'คณะผู้จัดทำ', 'ขอกราบ'
            ]
            
            for pattern in acknowledgement_content_patterns:
                if pattern in text:
                    return True
            return False

    def _is_keywords_section(self, text):
        """ตรวจสอบว่าเป็นส่วน Keywords หรือไม่"""
        keywords_patterns = [
            'keywords:', 'keyword:', 'คำสำคัญ:', 'คำสำคัญ :', 'key words:'
        ]
        text_lower = text.lower().strip()
        
        for pattern in keywords_patterns:
            if text_lower.startswith(pattern):
                return True
        return False

    def _is_table_of_contents_page_number(self, text):
        """ตรวจสอบว่าเป็นเลขหน้าในสารบัญหรือไม่"""
        text_clean = text.strip()
        
        # เลขหน้าที่อยู่ท้ายบรรทัดในสารบัญ
        # Header 2 คือหัวข้อใหญ่ ไม่ใช่ชื่อบท
        if re.match(r'^\d+', text_clean):
        
        # หัวข้อใหญ่ที่ขึ้นต้นด้วยตัวเลข (1. 2. 3.) ไม่ใช่ชื่อบท
            return False
        
        # ข้อความในกิตติกรรมประกาศที่ไม่ใช่ชื่อหัวข้อ - ไม่ใช่ชื่อบท
        if self._is_acknowledgement_content(text_clean):
            return False
        
        # 1. ตรวจสอบหัวข้อบท - บทที่ X
        if re.match(r'^บทที่\s*\d+', text_clean):
            return True
        
        # 2. ตรวจสอบชื่อบทที่รู้จัก (ไม่ขึ้นต้นด้วยตัวเลข)
        chapter_names = [
            'บทนำ', 'งานวิจัยและทฤษฎีที่เกี่ยวข้อง', 'การวิเคราะห์และออกแบบระบบ',
            'การพัฒนาระบบ', 'การทดสอบระบบ', 'สรุปและข้อเสนอแนะ',
            'เป้าหมายและขอบเขต', 'ปัญหาและขอบเขต'
        ]
        
        # ตรวจสอบว่าไม่ขึ้นต้นด้วยตัวเลข และมีชื่อบทที่รู้จัก
        if not re.match(r'^\d+\.', text_clean):
            for chapter_name in chapter_names:
                if chapter_name in text_clean:
                    return True
        
        # 3. ตรวจสอบหัวข้อหลักอื่นๆ ที่กำหนดใน standards (ไม่ขึ้นต้นด้วยตัวเลข)
        main_headings_extended = self.standards['main_headings'] + [
            'สารบัญตาราง', 'สารบัญภาพ', 'รายการสัญลักษณ์และคำย่อ',
            'บทที่ 1', 'บทที่ 2', 'บทที่ 3', 'บทที่ 4', 'บทที่ 5',
            'References', 'Bibliography', 'Appendix', 'สารบัญ(ต่อ)', 'สารบัญภาพ(ต่อ)', 'สารบัญตาราง(ต่อ)'
        ]
        
        # ตรวจสอบว่าไม่ขึ้นต้นด้วยตัวเลข และมีหัวข้อที่กำหนด
        if not re.match(r'^\d+\.', text_clean):
            for heading in main_headings_extended:
                if heading in text_clean:
                    return True
        
        return False

    def _is_major_heading(self, text, paragraph=None):
        """ตรวจสอบว่าเป็นหัวข้อใหญ่หรือไม่ - รูปแบบ 1. 2. 3. ขนาด 16pt Bold"""
        text_clean = text.strip()
        
        # ตรวจสอบ style ก่อน - Header 2 = หัวข้อใหญ่
        if paragraph and self._has_header2_style(paragraph):
            return True
        
        # ตรวจสอบรูปแบบหัวข้อใหญ่ 1. 2. 3. (ไม่ใช่ 1.1 หรือ 1.1.1)
        if re.match(r'^\d+\.\s+[^\d]', text_clean):  # เริ่มด้วยตัวเลข. ตามด้วยช่องว่างและไม่ใช่ตัวเลข
            return True
        
        return False

    def _is_sub_heading_level1(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับ 1 หรือไม่ - รูปแบบ 1.1 1.2 2.1 ขนาด 14pt ปกติ"""
        text_clean = text.strip()
        
        # รูปแบบ 1.1 1.2 2.1 (มีจุดสองจุด)
        if re.match(r'^\d+\.\d+\s', text_clean):
            return True
        
        return False

    def _is_sub_heading_level2(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับ 2 หรือไม่ - รูปแบบ 1.1.1 2.2.1 ขนาด 14pt ปกติ"""
        text_clean = text.strip()
        
        # รูปแบบ 1.1.1 2.2.1 (มีจุดสามจุด)
        if re.match(r'^\d+\.\d+\.\d+\s', text_clean):
            return True
        
        return False

    def _is_sub_heading_level3(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับ 3 หรือไม่ - รูปแบบ (1) (2) (3) ขนาด 14pt ปกติ"""
        text_clean = text.strip()
        
        # รูปแบบ (1) (2) (3)
        if re.match(r'^\(\d+\)\s', text_clean):
            return True
        
        return False

    def _is_any_sub_heading(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับใดๆ หรือไม่"""
        return (self._is_sub_heading_level1(text) or 
                self._is_sub_heading_level2(text) or 
                self._is_sub_heading_level3(text))

    def _has_header1_style(self, paragraph):
        """ตรวจสอบว่าเป็น Header 1 style หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'heading 1' in style_name or 'header 1' in style_name:
                    return True
        except:
            pass
        return False

    def _has_header2_style(self, paragraph):
        """ตรวจสอบว่าเป็น Header 2 style หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'heading 2' in style_name or 'header 2' in style_name:
                    return True
        except:
            pass
        return False

    def _check_alignment(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบการจัดแนว - ไม่ตรวจสอบการจัดกึ่งกลางของหน้าปก"""
        text = paragraph.text.strip()
        if not text:
            return
        
        # ข้ามการตรวจสอบการจัดแนวสำหรับหน้าปก
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
        
        # หัวข้อหลัก หัวข้อใหญ่ และหัวข้อย่อยควรชิดซ้าย - แก้ไขให้ครอบคลุม
        if ((self._is_main_heading(text, paragraph) or self._is_major_heading(text, paragraph) or self._is_any_sub_heading(text)) 
            and self.standards['main_heading_left']):
            if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER or paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'การจัดแนวไม่ถูกต้อง',
                    'description': 'หัวข้อควรจัดแนวชิดซ้าย',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, "หัวข้อควรจัดแนวชิดซ้าย")

    def _check_indentation(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบการเยื้อง - แก้ไขให้ไม่ตรวจสอบการเยื้องของ caption และส่วนพิเศษ"""
        text = paragraph.text.strip()
        if not text:
            return
        
        # ข้ามการตรวจสอบสำหรับหน้าปก
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
        
        # ข้ามการตรวจสอบสำหรับส่วนที่ไม่ต้องเยื้อง
        if self._should_skip_indent_check(text):
            return
        
        # ข้ามส่วน Keywords ใน Abstract
        if self._is_keywords_section(text):
            return
        
        # ข้ามเลขหน้าและคำว่า "ผู้จัดทำ" ในสารบัญ
        if self._is_table_of_contents_page_number(text):
            return
        
        # ข้ามหัวข้อทุกประเภท - แก้ไขให้ครอบคลุม
        if (self._is_main_heading(text, paragraph) or self._is_major_heading(text, paragraph) or self._is_any_sub_heading(text)):
            return
        
        # ข้ามการอ้างอิงตารางและรูปภาพ (ในเนื้อหา ไม่ใช่ caption)
        if self._contains_table_reference(text) or self._contains_picture_reference(text):
            return
        
        # ข้ามเนื้อหาในกิตติกรรมประกาศ
        if self._is_acknowledgement_content(text):
            return
        
        # เนื้อหาทั่วไปควรมีการเยื้องย่อหน้า
        if self.standards['paragraph_indent']:
            if paragraph.paragraph_format.first_line_indent is None:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'การเยื้องไม่ถูกต้อง',
                    'description': 'เนื้อหาควรมีการเยื้องย่อหน้า (ประมาณ 0.5 นิ้ว)',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, "เนื้อหาควรมีการเยื้องย่อหน้า (กด Tab หรือตั้งค่า First Line Indent 0.5 นิ้ว)")

    def _check_font_sizes(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบขนาดฟอนต์ - แก้ไขให้จำแนกประเภทหัวข้อถูกต้อง"""
        text = paragraph.text.strip()
        if not text:
            return
        
        # ข้ามหน้าปก (จะตรวจสอบแยกใน _check_cover_page_detailed)
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
            
        expected_size = None
        section_type = None
        
        # 1. ชื่อบท - 18pt Bold
        if self._is_main_heading(text, paragraph) or self._has_header1_style(paragraph):
            expected_size = 18
            section_type = "ชื่อบท"
        # 2. หัวข้อใหญ่ (1. 2. 3.) - 16pt Bold
        elif self._is_major_heading(text, paragraph):
            expected_size = 16
            section_type = "หัวข้อใหญ่"
        # 3. หัวข้อย่อยทุกระดับ (1.1, 1.1.1, (1)) - 14pt ปกติ
        elif self._is_any_sub_heading(text):
            expected_size = 14
            section_type = "หัวข้อย่อย"
        # 4. เนื้อหาทั่วไป - 14pt ปกติ
        else:
            expected_size = 14
            section_type = "เนื้อหา"
        
        # ตรวจสอบขนาดฟอนต์ในทุก run
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                actual_size = None
                
                try:
                    if run.font.size and hasattr(run.font.size, 'pt'):
                        actual_size = run.font.size.pt
                    elif paragraph.style and paragraph.style.font and paragraph.style.font.size:
                        if hasattr(paragraph.style.font.size, 'pt'):
                            actual_size = paragraph.style.font.size.pt
                except AttributeError:
                    continue
                
                # ให้ความผิดพลาด ±1pt สำหรับความยืดหยุ่น
                if actual_size and abs(actual_size - expected_size) > 1:
                    self.errors.append({
                        'paragraph': para_num,
                        'type': 'ขนาดฟอนต์ไม่ถูกต้อง',
                        'description': f'{section_type}ควรใช้ขนาด {expected_size}pt แต่พบขนาด {actual_size}pt',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, f"{section_type}ควรใช้ขนาด {expected_size}pt (ปัจจุบัน {actual_size}pt)")
                    return

    def _check_font_bold(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบความหนาของฟอนต์ - แก้ไขให้จำแนกประเภทหัวข้อถูกต้อง"""
        text = paragraph.text.strip()
        if not text or self._should_skip_bold_check(text):
            return
        
        # ข้ามหน้าปก (ตรวจสอบแยกแล้ว)
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
        
        should_be_bold = None
        section_type = None
        
        # 1. ชื่อบท - ควรเป็น Bold
        if self._is_main_heading(text, paragraph) or self._has_header1_style(paragraph):
            should_be_bold = True
            section_type = "ชื่อบท"
        # 2. หัวข้อใหญ่ (1. 2. 3.) - ควรเป็น Bold
        elif self._is_major_heading(text, paragraph):
            should_be_bold = True
            section_type = "หัวข้อใหญ่"
        # 3. หัวข้อย่อยทุกระดับ - ควรเป็นตัวปกติ (ไม่หนา)
        elif self._is_any_sub_heading(text):
            should_be_bold = False
            section_type = "หัวข้อย่อย"
        
        # ตรวจสอบความหนาเฉพาะเมื่อมีการกำหนด
        if should_be_bold is not None and section_type:
            has_bold_text = self._is_paragraph_bold(paragraph) or self._has_bold_style(paragraph)
            
            if should_be_bold and not has_bold_text:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'ความหนาฟอนต์ไม่ถูกต้อง',
                    'description': f'{section_type}ควรเป็นตัวหนา (Bold)',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, f"{section_type}ควรเป็นตัวหนา (Bold)")
            elif not should_be_bold and has_bold_text:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'ความหนาฟอนต์ไม่ถูกต้อง',
                    'description': f'{section_type}ควรเป็นตัวปกติ (ไม่หนา)',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, f"{section_type}ควรเป็นตัวปกติ (ไม่หนา)")

    def _is_sub_heading(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยหรือไม่ - รวมทุกระดับ (เพื่อความเข้ากันได้กับโค้ดเดิม)"""
        return self._is_any_sub_heading(text)

    def _is_project_title_improved(self, text, index, context):
        """ตรวจสอบว่าเป็นชื่อโครงงานหรือไม่ - ปรับปรุงใหม่ให้ทำงานได้ดีขึ้น"""
        text_clean = text.strip()
        
        # Debug: แสดงข้อมูลการตรวจสอบ
        # print(f"  Checking if project title: '{text_clean[:30]}...' (length: {len(text_clean)})")
        
        # เงื่อนไขพื้นฐาน: ความยาวมากกว่า 8 ตัวอักษร
        if len(text_clean) < 8:
            # print(f"    -> Too short: {len(text_clean)} < 8")
            return False
        
        # คำสำคัญที่บ่งบอกว่าไม่ใช่ชื่อโครงงาน
        non_title_keywords = [
            'โดย', 'อาจารย์', 'ที่ปรึกษา', 'ภาค', 'สาขา', 'รายงาน', 'วิทยาลัย', 
            'มหาวิทยาลัย', 'เดือน', 'พ.ศ', 'ปีการศึกษา', 'การศึกษา', 'cs/it', 
            'cs ', 'เอกสาร', 'ฉบับสมบูรณ์', 'ความก้าวหน้า', 'โครงงานฉบับสมบูรณ์'
        ]
        
        for keyword in non_title_keywords:
            if keyword.lower() in text_clean.lower():
                # print(f"    -> Contains non-title keyword: {keyword}")
                return False
        
        # ไม่ใช่รหัสนักศึกษา
        if re.search(r'\d{9,10}-\d', text_clean):
            # print(f"    -> Contains student ID pattern")
            return False
        
        # ไม่เริ่มต้นด้วยตัวเลขหรือสัญลักษณ์พิเศษ (ยกเว้นวงเล็บ)
        if re.match(r'^[\d\W&&[^(]]', text_clean):
            # print(f"    -> Starts with number or special character")
            return False
        
        # ตรวจสอบตำแหน่งใน context
        if context.get('project_type_index', -1) >= 0 and context.get('author_section_start', -1) >= 0:
            # ถ้าอยู่ระหว่างประเภทโครงงานและส่วนผู้แต่ง
            if context['project_type_index'] < index < context['author_section_start']:
                # print(f"    -> Position check passed: {context['project_type_index']} < {index} < {context['author_section_start']}")
                return True
        
        # ตรวจสอบรูปแบบภาษาอังกฤษที่อาจเป็นชื่อโครงงาน
        if re.search(r'[A-Z][a-z]+.*[A-Z]', text_clean):  # มีการผสมตัวใหญ่-เล็กแบบชื่อเรื่อง
            # print(f"    -> English title pattern detected")
            return True
        
        # ตรวจสอบรูปแบบภาษาไทยที่อาจเป็นชื่อโครงงาน
        if re.search(r'[ก-๙].*ระบบ|[ก-๙].*การ|[ก-๙].*โครงงาน', text_clean):
            # print(f"    -> Thai title pattern detected")
            return True
        
        # เพิ่มการตรวจสอบความยาวที่เหมาะสม (ชื่อโครงงานมักจะยาว)
        if len(text_clean) > 15 and not any(char.isdigit() for char in text_clean[:5]):
            # print(f"    -> Long text without numbers at start")
            return True
        
        # print(f"    -> No conditions matched")
        return False

    def _debug_cover_structure(self, cover_paragraphs, cover_context):
        """ฟังก์ชัน debug สำหรับดูโครงสร้างหน้าปก"""
        print("=== Debug Cover Structure ===")
        for i, para in enumerate(cover_paragraphs):
            text = para.text.strip()
            if text:
                print(f"Index {i}: '{text[:50]}...' (length: {len(text)})")
                if i in cover_context['project_title_candidates']:
                    print(f"  -> ระบุเป็นชื่อโครงงาน candidate")
                if self._is_project_title_improved(text, i, cover_context):
                    print(f"  -> ผ่านการตรวจสอบ _is_project_title_improved")
        
        print(f"Project type index: {cover_context['project_type_index']}")
        print(f"Author section start: {cover_context['author_section_start']}")
        print(f"Project title candidates: {cover_context['project_title_candidates']}")
        print("=============================")

    def _get_expected_cover_font_size_with_context(self, text, index, context):
        """กำหนดขนาดฟอนต์หน้าปกโดยใช้ context - แก้ไขให้ตรงกับมาตรฐาน PDF"""
        text_clean = text.strip()
        text_lower = text_clean.lower()
        
        # Debug: แสดงข้อมูลการตรวจสอบ
        # print(f"Checking font size for index {index}: '{text_clean[:30]}...'")
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม = 26pt
        if re.search(r'cs\s*\d{4}', text_lower) or 'cs/it' in text_lower:
            # print(f"  -> CS code detected: 26pt")
            return 26
            
        # 2. ชื่อเล่ม = 20pt
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า']):
            # print(f"  -> Project type detected: 20pt")
            return 20
            
        # 3. ชื่อโครงงาน (ใช้ context) = 20pt - ปรับปรุงลำดับการตรวจสอบ
        elif (index in context['project_title_candidates'] or 
            self._is_project_title_improved(text_clean, index, context)):
            # print(f"  -> Project title detected: 20pt")
            return 20
            
        # 4. คำว่า "โดย" และชื่อผู้ทำ = 18pt
        elif ('โดย' in text_clean or re.search(r'\d{9,10}-\d', text_clean) or 
            (context['author_section_start'] > 0 and 
            context['author_section_start'] <= index < context.get('advisor_section_start', float('inf')))):
            # print(f"  -> Author section detected: 18pt")
            return 18
            
        # 5. อาจารย์ที่ปรึกษา = 18pt
        elif (any(keyword in text_clean for keyword in ['อาจารย์', 'ที่ปรึกษา']) or
            (context['advisor_section_start'] > 0 and 
            context['advisor_section_start'] <= index < context.get('bottom_section_start', float('inf')))):
            # print(f"  -> Advisor section detected: 18pt")
            return 18
            
        # 6. ส่วนด้านล่าง = 16pt
        elif (context['bottom_section_start'] > 0 and index >= context['bottom_section_start']) or \
            any(keyword in text_clean for keyword in ['สาขา', 'วิทยาลัย', 'มหาวิทยาลัย', 'ภาคเรียน', 'เดือน', 'พ.ศ', 'รายงานนี้', 'การศึกษา']):
            # print(f"  -> Bottom section detected: 16pt")
            return 16
        
        # print(f"  -> No specific rule matched")
        return None

    def _analyze_cover_structure(self, cover_paragraphs):
        """วิเคราะห์โครงสร้างหน้าปกเพื่อระบุตำแหน่งของแต่ละส่วน - ปรับปรุงให้แม่นยำขึ้น"""
        context = {
            'cs_code_index': -1,
            'project_type_index': -1,
            'author_section_start': -1,
            'advisor_section_start': -1,
            'bottom_section_start': -1,
            'project_title_candidates': []
        }
        
        for i, para in enumerate(cover_paragraphs):
            text = para.text.strip()
            text_lower = text.lower()
            
            # หารหัสสาขาและปี
            if re.search(r'cs\s*\d{4}', text_lower) or 'cs/it' in text_lower:
                context['cs_code_index'] = i
            
            # หาประเภทโครงงาน - ตรวจสอบให้แม่นยำขึ้น
            elif ('โครงงานฉบับสมบูรณ์' in text or 'เอกสารโครงงาน' in text or 
                'รายงานความก้าวหน้า' in text):
                context['project_type_index'] = i
            
            # หาส่วนผู้แต่ง - ปรับให้ยืดหยุ่นขึ้น
            elif ('โดย' in text and len(text.strip()) <= 10):  # คำว่า "โดย" อย่างเดียวหรือใกล้เคียง
                context['author_section_start'] = i
            
            # หาส่วนอาจารย์ที่ปรึกษา
            elif 'อาจารย์ที่ปรึกษา' in text:
                context['advisor_section_start'] = i
            
            # หาส่วนข้อมูลด้านล่าง
            elif any(keyword in text for keyword in ['รายงานนี้เป็นส่วนหนึ่ง', 'สาขาวิชา', 'วิทยาลัย', 'มหาวิทยาลัย']):
                if context['bottom_section_start'] == -1:
                    context['bottom_section_start'] = i
        
        # ระบุชื่อโครงงานจากตำแหน่ง - ปรับปรุงการค้นหาให้แม่นยำขึ้น
        if context['project_type_index'] >= 0 and context['author_section_start'] >= 0:
            project_start = context['project_type_index'] + 1
            project_end = context['author_section_start']
            
            for i in range(project_start, project_end):
                if i < len(cover_paragraphs):
                    text = cover_paragraphs[i].text.strip()
                    # ปรับเงื่อนไข: ให้ยืดหยุ่นขึ้นในการระบุชื่อโครงงาน
                    if (len(text) > 5 and  # ลดจาก 10 เป็น 5 เพื่อให้ครอบคลุมมากขึ้น
                        not any(keyword in text.lower() for keyword in ['โดย', 'อาจารย์', 'cs/it', 'cs ', 'รหัส']) and
                        not re.search(r'\d{9,10}-\d', text) and  # ไม่ใช่รหัสนักศึกษา
                        not text.isdigit() and  # ไม่ใช่ตัวเลขเพียงอย่างเดียว
                        text not in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน']):  # ไม่ใช่ชื่อประเภทเอกสาร
                        context['project_title_candidates'].append(i)
        
        return context
    
    def _check_cover_page_detailed(self, doc):
        """ตรวจสอบหน้าปกตามโครงสร้างใหม่ - เพิ่ม debug และปรับปรุงการระบุชื่อโครงงาน"""
        cover_end_index = self._find_cover_end_index(doc)
        cover_paragraphs = [p for p in doc.paragraphs[:cover_end_index] if p.text.strip()]
        
        # สร้าง context สำหรับการวิเคราะห์ตำแหน่ง
        cover_context = self._analyze_cover_structure(cover_paragraphs)
        
        # Debug: แสดงโครงสร้างหน้าปก (comment out ในการใช้งานจริง)
        # self._debug_cover_structure(cover_paragraphs, cover_context)
        
        # ตรวจสอบเฉพาะขนาดฟอนต์และความหนาของหน้าปก
        for i, para in enumerate(cover_paragraphs):
            text = para.text.strip()
            if not text:
                continue
            
            # ใช้ context ในการกำหนดขนาดฟอนต์
            expected_size = self._get_expected_cover_font_size_with_context(text, i, cover_context)
            if expected_size:
                actual_size = self._get_paragraph_font_size(para)
                if actual_size and abs(actual_size - expected_size) > 2:  # ให้ความผิดพลาด ±2pt
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'ขนาดฟอนต์หน้าปกไม่ถูกต้อง',
                        'description': f'ข้อความ "{text[:30]}..." ในหน้าปกควรใช้ขนาด {expected_size}pt แต่พบ {actual_size}pt',
                        'paragraph_obj': para
                    })
                    self._highlight_paragraph(para)
                    self._add_comment_to_paragraph(para, f'ขนาดฟอนต์ควร {expected_size}pt')
            
            # ตรวจสอบความหนาเฉพาะส่วนที่ควรเป็น Bold
            expected_bold = self._should_cover_text_be_bold_with_context(text, i, cover_context)
            if expected_bold is not None:
                is_bold = self._is_paragraph_bold(para) or self._has_bold_style(para)
                if expected_bold and not is_bold:
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'ความหนาฟอนต์หน้าปกไม่ถูกต้อง',
                        'description': f'ข้อความ "{text[:30]}..." ในหน้าปกควรเป็นตัวหนา',
                        'paragraph_obj': para
                    })
                    self._highlight_paragraph(para)
                    self._add_comment_to_paragraph(para, 'ข้อความนี้ในหน้าปกควรเป็นตัวหนา')
                elif expected_bold == False and is_bold:
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'ความหนาฟอนต์หน้าปกไม่ถูกต้อง',
                        'description': f'ข้อความ "{text[:30]}..." ในหน้าปกควรเป็นตัวปกติ (ไม่หนา)',
                        'paragraph_obj': para
                    })
                    self._highlight_paragraph(para)
                    self._add_comment_to_paragraph(para, 'ข้อความนี้ในหน้าปกควรเป็นตัวปกติ (ไม่หนา)')

    def _should_cover_text_be_bold_with_context(self, text, index, context):
        """กำหนดว่าข้อความในหน้าปกควรเป็น Bold หรือไม่ โดยใช้ context"""
        text_clean = text.strip()
        text_lower = text_clean.lower()
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม = หนา
        if re.search(r'cs\s*\d{4}', text_lower) or 'cs/it' in text_lower:
            return True
            
        # 2. ชื่อเล่ม = หนา
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า']):
            return True
            
        # 3. ชื่อโครงงาน = ปกติ (ไม่หนา) - ปรับปรุงลำดับการตรวจสอบ
        elif (index in context['project_title_candidates'] or 
            self._is_project_title_improved(text_clean, index, context)):
            return False
            
        # 4. ส่วนอื่นๆ = ปกติ (ไม่หนา)
        else:
            return False
    
    def _get_expected_cover_font_size_from_pdf(self, text):
        """กำหนดขนาดฟอนต์หน้าปกตามโครงสร้างที่อธิบาย - ปรับปรุงใหม่"""
        text_clean = text.strip()   
        text_lower = text_clean.lower()
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม เช่น CS 2567/CS-59 = 26pt หนา
        if re.search(r'cs\s*\d{4}', text_lower) or re.search(r'cs/it.*\d{4}', text_lower) or 'cs/it' in text_lower:
            return 26
            
        # 2. ชื่อเล่ม เช่น โครงงานฉบับสมบูรณ์, รายงานความก้าวหน้า = 20pt หนา  
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า', 'โครงงาน']):
            return 20
            
        # 3. ชื่อเรื่องโครงงานภาษาไทยและอังกฤษ = 20pt ปกติ
        # ปรับปรุง: ใช้วิธีการระบุตำแหน่งที่แม่นยำขึ้น
        elif self._is_project_title(text_clean):
            return 20
            
        # 4. คำว่า "โดย" และชื่อผู้ทำโครงงาน = 18pt ปกติ
        elif 'โดย' in text_clean or re.search(r'\d{9,10}-\d', text_clean):
            return 18
            
        # 5. ชื่ออาจารย์ที่ปรึกษา = 18pt ปกติ  
        elif any(keyword in text_clean for keyword in ['อาจารย์', 'ที่ปรึกษา']):
            return 18
            
        # 6. ส่วนด้านล่าง (ข้อมูลสาขา, ภาคเรียน, เดือน, ปี) = 16pt ปกติ
        elif any(keyword in text_clean for keyword in ['สาขา', 'วิทยาลัย', 'มหาวิทยาลัย', 'ภาคเรียน', 'เดือน', 'พ.ศ', 'รายงานนี้', 'การศึกษา', 'ปีการศึกษา']):
            return 16
        
        return None
        
    def _should_cover_text_be_bold_from_pdf(self, text):
        """กำหนดว่าข้อความในหน้าปกควรเป็น Bold หรือไม่ - ปรับปรุงใหม่"""
        text_clean = text.strip()
        text_lower = text_clean.lower()
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม เช่น CS 2567/CS-59 = หนา
        if re.search(r'cs\s*\d{4}', text_lower) or re.search(r'cs/it.*\d{4}', text_lower) or 'cs/it' in text_lower:
            return True
            
        # 2. ชื่อเล่ม เช่น โครงงานฉบับสมบูรณ์ = หนา
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า']):
            return True
            
        # 3. ชื่อเรื่องโครงงาน = ปกติ (ไม่หนา)
        elif self._is_project_title(text_clean):
            return False
            
        # 4. ชื่อผู้ทำโครงงาน = ปกติ
        elif 'โดย' in text_clean or re.search(r'\d{9,10}-\d', text_clean):
            return False
            
        # 5. ชื่ออาจารย์ที่ปรึกษา = ปกติ
        elif any(keyword in text_clean for keyword in ['อาจารย์', 'ที่ปรึกษา']):
            return False
            
        # 6. ส่วนด้านล่าง = ปกติ
        elif any(keyword in text_clean for keyword in ['สาขา', 'วิทยาลัย', 'มหาวิทยาลัย', 'ภาคเรียน', 'เดือน', 'พ.ศ', 'รายงานนี้', 'การศึกษา', 'ปีการศึกษา']):
            return False
        
        return None
    
    def _is_project_title(self, text):
        """ตรวจสอบว่าเป็นชื่อโครงงานหรือไม่"""
        text_clean = text.strip()
        
        # เงื่อนไขสำหรับระบุชื่อโครงงาน:
        # 1. ความยาวมากกว่า 15 ตัวอักษร (ชื่อโครงงานมักจะยาว)
        # 2. ไม่มีคำสำคัญที่บ่งบอกว่าไม่ใช่ชื่อโครงงาน
        # 3. ไม่ใช่รหัสนักศึกษา
        # 4. ไม่ใช่ข้อมูลส่วนอื่นๆ
        
        if len(text_clean) < 15:
            return False
        
        # คำสำคัญที่บ่งบอกว่าไม่ใช่ชื่อโครงงาน
        non_title_keywords = [
            'โดย', 'อาจารย์', 'ที่ปรึกษา', 'ภาค', 'สาขา', 'รายงาน', 'วิทยาลัย', 
            'มหาวิทยาลัย', 'เดือน', 'พ.ศ', 'ปีการศึกษา', 'การศึกษา', 'cs/it', 
            'cs ', 'โครงงาน', 'เอกสาร', 'ฉบับสมบูรณ์', 'ความก้าวหน้า'
        ]
        
        for keyword in non_title_keywords:
            if keyword.lower() in text_clean.lower():
                return False
        
        # ไม่ใช่รหัสนักศึกษา
        if re.search(r'\d{9,10}-\d', text_clean):
            return False
        
        # ไม่เริ่มต้นด้วยตัวเลขหรือสัญลักษณ์พิเศษ
        if re.match(r'^[\d\W]', text_clean):
            return False
        
        return True

    def _find_cover_end_index(self, doc):
        """หาจุดสิ้นสุดของหน้าปก - ปรับปรุงให้รองรับหน้าปก 2 หน้าที่ไม่มีเลขหน้า"""
        cover_end_index = 0
        found_page_number = False
        
        for i, para in enumerate(doc.paragraphs[:60]):  # เพิ่มขอบเขตการค้นหา
            text = para.text.strip().lower()
            
            # ตรวจสอบเลขหน้า - หากเจอเลขหน้าแสดงว่าจบหน้าปกแล้ว
            if re.search(r'^\s*\d+\s*$', para.text.strip()) and i > 10:  # เลขหน้าที่ปรากฏหลังจากเนื้อหาหน้าปกบ้าง
                found_page_number = True
                cover_end_index = i
                break
            
            # คำสำคัญที่บ่งชี้ว่าจบหน้าปกแล้ว
            end_cover_keywords = [
                'บทคัดย่อ', 'abstract', 'สารบัญ', 'บทที่', 'กิตติกรรม', 
                'acknowledgement', 'table of contents', 'contents', 
                'ประกาศนียบัตร', 'certificate', 'คำนำ', 'preface'
            ]
            
            if any(keyword in text for keyword in end_cover_keywords):
                cover_end_index = i
                break
        
        # หากไม่เจอคำสำคัญใดๆ ให้ใช้ค่า default สำหรับหน้าปก 2 หน้า
        if cover_end_index == 0:
            cover_end_index = 40  # ประมาณ 2 หน้า
        
        return cover_end_index

    def _check_references(self, doc):
        """ตรวจสอบการอ้างอิงตารางและรูปภาพ - แก้ไขให้ตรวจสอบจาก caption style"""
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            
            # ตรวจสอบว่าเป็น caption หรือไม่จาก style
            is_caption = self._is_caption_style(paragraph)
            
            if is_caption:
                # ถ้าเป็น caption แล้วข้ามการตรวจสอบการเยื้อง
                continue
            
            # ตรวจสอบการอ้างอิงตารางในเนื้อหา (ไม่ใช่ caption)
            if self._contains_table_reference(text) and not is_caption:
                if not self._is_valid_table_reference_format(text):
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'รูปแบบการอ้างอิงตารางไม่ถูกต้อง',
                        'description': 'การอ้างอิงตารางควรเป็น "ตารางที่ X" หรือ "ตารางที่ X.X"',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, 'การอ้างอิงตารางควรเป็น "ตารางที่ X" หรือ "ตารางที่ X.X"')
            
            # ตรวจสอบการอ้างอิงรูปภาพในเนื้อหา (ไม่ใช่ caption)
            if self._contains_picture_reference(text) and not is_caption:
                if not self._is_valid_picture_reference_format(text):
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'รูปแบบการอ้างอิงรูปภาพไม่ถูกต้อง',
                        'description': 'การอ้างอิงรูปภาพควรเป็น "ภาพที่ X" หรือ "ภาพที่ X.X"',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, 'การอ้างอิงรูปภาพควรเป็น "ภาพที่ X" หรือ "ภาพที่ X.X"')

    def _is_caption_style(self, paragraph):
        """ตรวจสอบว่าเป็น caption style หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'caption' in style_name:
                    return True
        except:
            pass
        return False

    def _contains_table_reference(self, text):
        """ตรวจสอบว่าข้อความมีการอ้างอิงตารางหรือไม่"""
        patterns = [
            r'ตารางที่\s*\d+',
            r'ตารางที่\s*\d+\.\d+',
            r'table\s*\d+',
            r'Table\s*\d+',
        ]
        for pattern in patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return True
        return False
    
    def _contains_picture_reference(self, text):
        """ตรวจสอบว่าข้อความมีการอ้างอิงรูปภาพหรือไม่"""
        patterns = [
            r'ภาพที่\s*\d+',
            r'ภาพที่\s*\d+\.\d+',
            r'รูปที่\s*\d+',
            r'รูปที่\s*\d+\.\d+',
            r'figure\s*\d+',
            r'Figure\s*\d+',
        ]
        for pattern in patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return True
        return False
    
    def _is_valid_table_reference_format(self, text):
        """ตรวจสอบรูปแบบการอ้างอิงตารางว่าถูกต้องหรือไม่"""
        valid_patterns = [
            r'^ตารางที่\s*\d+\s',
            r'^ตารางที่\s*\d+\.\d+\s',
            r'^ตารางที่\s*\d+',
            r'^ตารางที่\s*\d+\.\d+',
        ]
        for pattern in valid_patterns:
            if re.search(pattern, text):
                return True
        return False
    
    def _is_valid_picture_reference_format(self, text):
        """ตรวจสอบรูปแบบการอ้างอิงรูปภาพว่าถูกต้องหรือไม่"""
        valid_patterns = [
            r'^ภาพที่\s*\d+\s',
            r'^ภาพที่\s*\d+\.\d+\s',
            r'^ภาพที่\s*\d+',
            r'^ภาพที่\s*\d+\.\d+',
        ]
        for pattern in valid_patterns:
            if re.search(pattern, text):
                return True
        return False

    def check_document_object(self, doc):
        """ตรวจสอบ Document object ที่ผ่านเข้ามา"""
        self.errors = []
        self.tables = []
        self.pictures = []
        
        # ตรวจหาตารางและรูปภาพ
        self._extract_tables_and_pictures(doc)
        
        # ตรวจสอบหน้าปก
        self._check_cover_page_detailed(doc)
        
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():
                # ข้ามหน้าปกในการตรวจสอบปกติ
                if not self._is_cover_page_section(paragraph.text, i):
                    self._check_paragraph(paragraph, i + 1, i)
        
        # ตรวจสอบการอ้างอิงตารางและรูปภาพ
        self._check_references(doc)
        
        return self.errors
    
    def _get_paragraph_font_size(self, paragraph):
        """ดึงขนาดฟอนต์จากย่อหน้า"""
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                if run.font.size and hasattr(run.font.size, 'pt'):
                    return run.font.size.pt
        
        try:
            if paragraph.style and paragraph.style.font and paragraph.style.font.size:
                if hasattr(paragraph.style.font.size, 'pt'):
                    return paragraph.style.font.size.pt
        except:
            pass
            
        return None
    
    def _is_paragraph_bold(self, paragraph):
        """ตรวจสอบว่าย่อหน้าเป็นตัวหนาหรือไม่"""
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                if run.font.bold:
                    return True
        return False

    def _has_bold_style(self, paragraph):
        """ตรวจสอบว่าย่อหน้ามี style ที่เป็น bold หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'font'):
                if paragraph.style.font.bold:
                    return True
        except:
            pass
        return False
    
    def _check_paragraph(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบย่อหน้าแต่ละย่อหน้า"""
        self._check_thai_font(paragraph, para_num, paragraph_index)
        self._check_font_sizes(paragraph, para_num, paragraph_index)
        self._check_font_bold(paragraph, para_num, paragraph_index)
        self._check_alignment(paragraph, para_num, paragraph_index)
        self._check_indentation(paragraph, para_num, paragraph_index)
        self._check_spacing(paragraph, para_num, paragraph_index)
        self._check_heading_numbering(paragraph, para_num, paragraph_index)
    
    def _is_cover_page_section(self, text, paragraph_index):
        """ตรวจสอบว่าเป็นส่วนของหน้าปกหรือไม่ - แก้ไขให้รองรับหน้าปก 2 หน้า"""
        # เพิ่มขอบเขตการตรวจสอบหน้าปกให้มากขึ้น (รองรับ 2 หน้า)
        if paragraph_index > 50:  # เพิ่มจาก 25 เป็น 50
            return False
        
        # ตรวจสอบคำสำคัญที่บ่งบอกว่าไม่ใช่หน้าปก
        non_cover_keywords = ['บทคัดย่อ', 'abstract', 'สารบัญ', 'บทที่', 'กิตติกรรม', 'acknowledgement', 'table of contents', 'contents', 'ประกาศนียบัตร', 'certificate']
        for keyword in non_cover_keywords:
            if keyword.lower() in text.lower():
                return False
        
        # ตรวจสอบคำสำคัญที่บ่งบอกว่าเป็นหน้าปก
        cover_keywords_from_standards = self.standards.get('cover_sections', [])
        for keyword in cover_keywords_from_standards:
            if keyword in text:
                return True
        
        # ตรวจสอบรหัสนักศึกษา (รูปแบบ ตัวเลข 9-10 หลัก ตามด้วย - และตัวเลข 1 หลัก)
        if re.search(r'\d{9,10}-\d', text):
            return True
        
        # ตรวจสอบคำสำคัญเพิ่มเติมสำหรับหน้าปก
        additional_cover_keywords = [
            'cs/it', 'cs ', 'ภาคเรียนที่', 'ปีการศึกษา', 'เดือน', 'พ.ศ.', 'ครั้งที่', 
            'วิทยาลัยการคอมพิวเตอร์', 'วิทยาลัยคอมพิวเตอร์', 'มหาวิทยาลัยขอนแก่น', 
            'โครงงาน', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า', 'โดย', 'อาจารย์ที่ปรึกษา',
            'รายงานนี้เป็นส่วนหนึ่ง', 'การศึกษาวิชา'
        ]
        
        for keyword in additional_cover_keywords:
            if keyword.lower() in text.lower():
                return True
                
        return False
    
    def _has_header_style(self, paragraph):
        """ตรวจสอบว่าย่อหน้ามี Header style หรือไม่ - แก้ไขให้ตรวจสอบ Header 1 เป็นชื่อบท"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'heading' in style_name or 'header' in style_name:
                    return True
        except:
            pass
        return False
    
    def _check_spacing(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบระยะห่างบรรทัด"""
        if not paragraph.text.strip():
            return
            
        line_spacing = paragraph.paragraph_format.line_spacing
        if line_spacing and line_spacing != 1.0:
            if isinstance(line_spacing, (int, float)) and line_spacing > 1.1:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'ระยะห่างบรรทัดไม่ถูกต้อง',
                    'description': f'ควรใช้ระยะห่างบรรทัด 1 เท่า (Single) แต่พบ {line_spacing}',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, f"ควรใช้ระยะห่างบรรทัด 1 เท่า (Single) - ปัจจุบัน {line_spacing}")
    
    def _check_heading_numbering(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบการใช้เลขหัวข้อ"""
        text = paragraph.text.strip()
        
        if re.match(r'^\d+(\.\d+)*\s*$', text):
            self.errors.append({
                'paragraph': para_num,
                'type': 'รูปแบบหัวข้อไม่ถูกต้อง',
                'description': 'หัวข้อควรมีเนื้อหาตามหลังเลขหัวข้อ',
                'paragraph_obj': paragraph
            })
            self._highlight_paragraph(paragraph)
            self._add_comment_to_paragraph(paragraph, "หัวข้อควรมีเนื้อหาตามหลังเลขหัวข้อ")
    
    def _should_skip_indent_check(self, text):
        """ตรวจสอบว่าควรข้ามการตรวจสอบการเยื้องหรือไม่"""
        # เพิ่มรายการที่ไม่ต้องตรวจสอบการเยื้อง
        additional_skip_indent = [
            'keywords:', 'keyword:', 'คำสำคัญ:', 'คำสำคัญ :',
            'หน้า', 'ผู้จัดทำ', 'page', 
            'ตารางที่', 'ภาพที่', 'รูปที่', 'table', 'figure'
        ]
        
        text_lower = text.lower()
        
        # ตรวจสอบรายการเดิม
        for section in self.standards['no_indent_sections']:
            if section in text:
                return True
        
        # ตรวจสอบรายการเพิ่มเติม
        for item in additional_skip_indent:
            if item in text_lower:
                return True
                
        return False
    
    def _should_skip_font_check(self, text):
        """ตรวจสอบว่าควรข้ามการตรวจสอบฟอนต์หรือไม่"""
        for section in self.standards['skip_font_check']:
            if section in text:
                return True
        return False
    
    def _should_skip_bold_check(self, text):
        """ตรวจสอบว่าควรข้ามการตรวจสอบความหนาหรือไม่"""
        # เพิ่มรายการที่ไม่ต้องตรวจสอบความหนา
        additional_skip_bold = [
            'keywords:', 'keyword:', 'คำสำคัญ:', 'คำสำคัญ :',
            'ขอขอบพระคุณ', 'ขอขอบคุณ', 'คุณพ่อ', 'คุณแม่',
            'ผู้จัดทำ', 'หน้า'
        ]
        
        text_lower = text.lower()
        
        # ตรวจสอบรายการเดิม
        for section in self.standards['skip_bold_check']:
            if section in text:
                return True
        
        # ตรวจสอบรายการเพิ่มเติม
        for item in additional_skip_bold:
            if item in text_lower:
                return True
        
        # ข้ามเนื้อหาในกิตติกรรมประกาศ
        if self._is_acknowledgement_content(text):
            return True
                
        return False
    
    def _add_comment_to_paragraph(self, paragraph, comment_text):
        """เพิ่ม comment ที่ท้ายย่อหน้า"""
        if "[❌ ข้อผิดพลาด:" in paragraph.text:
            return
        
        new_run = paragraph.add_run(f" [❌ ข้อผิดพลาด: {comment_text}]")
        new_run.font.color.rgb = RGBColor(255, 0, 0)
        new_run.font.bold = True
        new_run.font.size = Pt(12)
    
    def _highlight_paragraph(self, paragraph):
        """ทำ highlight ย่อหน้า"""
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                run.font.highlight_color = 7
    
    def _check_thai_font(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบฟอนต์ภาษาไทย"""
        text = paragraph.text.strip()
        if not text or self._should_skip_font_check(text):
            return
        
        has_thai = any(ord(char) >= 0x0E00 and ord(char) <= 0x0E7F for char in text)
        if not has_thai:
            return
        
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                font_name = run.font.name
                if font_name and font_name != self.standards['theme_font']:
                    self.errors.append({
                        'paragraph': para_num,
                        'type': 'ฟอนต์ไม่ถูกต้อง',
                        'description': f'ข้อความภาษาไทยควรใช้ฟอนต์ {self.standards["theme_font"]} (พบ {font_name})',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, f"ข้อความภาษาไทยควรใช้ฟอนต์ {self.standards['theme_font']} (ปัจจุบันใช้ {font_name})")
                    return
    
    def get_error_statistics(self):
        """สร้างสถิติข้อผิดพลาด"""
        error_types = {}
        for error in self.errors:
            error_type = error['type']
            error_types[error_type] = error_types.get(error_type, 0) + 1
        return error_types
        # คำว่า "หน้า" หรือ "ผู้จัดทำ" ที่อยู่ในสารบัญ
        if text_clean in ['หน้า', 'ผู้จัดทำ', 'Page']:
            return True
            
        return False

    def _is_section2_content(self, text, paragraph_index):
        """ตรวจสอบว่าเป็นเนื้อหาใน Section 2 หรือไม่"""
        # Section 2 ประกอบด้วย: บทคัดย่อ, Abstract, กิตติกรรมประกาศ, สารบัญ
        section2_indicators = [
            'บทคัดย่อ', 'abstract', 'กิตติกรรมประกาศ', 'acknowledgement',
            'สารบัญ', 'table of contents', 'สารบัญภาพ', 'สารบัญตาราง',
            'list of figures', 'list of tables'
        ]
        
        text_lower = text.lower()
        for indicator in section2_indicators:
            if indicator in text_lower:
                return True
        
        # ตรวจสอบตำแหน่งโดยประมาณ (หลังหน้าปกแต่ก่อนบทที่ 1)
        if 40 <= paragraph_index <= 200:  # ประมาณการตำแหน่ง Section 2
            return True
            
        return False  # Header 2 คือหัวข้อใหญ่ ไม่ใช่ชื่อบท
        
        # หัวข้อใหญ่ที่ขึ้นต้นด้วยตัวเลข (1. 2. 3.) ไม่ใช่ชื่อบท
        if re.match(r'^\d+\.\s+[^\d]', text_clean):
            return False
        
        # ข้อความในกิตติกรรมประกาศที่ไม่ใช่ชื่อหัวข้อ - ไม่ใช่ชื่อบท
        if self._is_acknowledgement_content(text_clean):
            return False
        
        # 1. ตรวจสอบหัวข้อบท - บทที่ X
        if re.match(r'^บทที่\s*\d+', text_clean):
            return True
        
        # 2. ตรวจสอบชื่อบทที่รู้จัก (ไม่ขึ้นต้นด้วยตัวเลข)
        chapter_names = [
            'บทนำ', 'งานวิจัยและทฤษฎีที่เกี่ยวข้อง', 'การวิเคราะห์และออกแบบระบบ',
            'การพัฒนาระบบ', 'การทดสอบระบบ', 'สรุปและข้อเสนอแนะ',
            'เป้าหมายและขอบเขต', 'ปัญหาและขอบเขต'
        ]
        
        # ตรวจสอบว่าไม่ขึ้นต้นด้วยตัวเลข และมีชื่อบทที่รู้จัก
        if not re.match(r'^\d+\.', text_clean):
            for chapter_name in chapter_names:
                if chapter_name in text_clean:
                    return True
        
        # 3. ตรวจสอบหัวข้อหลักอื่นๆ ที่กำหนดใน standards (ไม่ขึ้นต้นด้วยตัวเลข)
        main_headings_extended = self.standards['main_headings'] + [
            'สารบัญตาราง', 'สารบัญภาพ', 'รายการสัญลักษณ์และคำย่อ',
            'บทที่ 1', 'บทที่ 2', 'บทที่ 3', 'บทที่ 4', 'บทที่ 5',
            'References', 'Bibliography', 'Appendix', 'สารบัญ(ต่อ)', 'สารบัญภาพ(ต่อ)', 'สารบัญตาราง(ต่อ)'
        ]
        
        # ตรวจสอบว่าไม่ขึ้นต้นด้วยตัวเลข และมีหัวข้อที่กำหนด
        if not re.match(r'^\d+\.', text_clean):
            for heading in main_headings_extended:
                if heading in text_clean:
                    return True
        
        return False

    def _is_major_heading(self, text, paragraph=None):
        """ตรวจสอบว่าเป็นหัวข้อใหญ่หรือไม่ - รูปแบบ 1. 2. 3. ขนาด 16pt Bold"""
        text_clean = text.strip()
        
        # ตรวจสอบ style ก่อน - Header 2 = หัวข้อใหญ่
        if paragraph and self._has_header2_style(paragraph):
            return True
        
        # ตรวจสอบรูปแบบหัวข้อใหญ่ 1. 2. 3. (ไม่ใช่ 1.1 หรือ 1.1.1)
        if re.match(r'^\d+\.\s+[^\d]', text_clean):  # เริ่มด้วยตัวเลข. ตามด้วยช่องว่างและไม่ใช่ตัวเลข
            return True
        
        return False

    def _is_sub_heading_level1(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับ 1 หรือไม่ - รูปแบบ 1.1 1.2 2.1 ขนาด 14pt ปกติ"""
        text_clean = text.strip()
        
        # รูปแบบ 1.1 1.2 2.1 (มีจุดสองจุด)
        if re.match(r'^\d+\.\d+\s', text_clean):
            return True
        
        return False

    def _is_sub_heading_level2(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับ 2 หรือไม่ - รูปแบบ 1.1.1 2.2.1 ขนาด 14pt ปกติ"""
        text_clean = text.strip()
        
        # รูปแบบ 1.1.1 2.2.1 (มีจุดสามจุด)
        if re.match(r'^\d+\.\d+\.\d+\s', text_clean):
            return True
        
        return False

    def _is_sub_heading_level3(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับ 3 หรือไม่ - รูปแบบ (1) (2) (3) ขนาด 14pt ปกติ"""
        text_clean = text.strip()
        
        # รูปแบบ (1) (2) (3)
        if re.match(r'^\(\d+\)\s', text_clean):
            return True
        
        return False

    def _is_any_sub_heading(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยระดับใดๆ หรือไม่"""
        return (self._is_sub_heading_level1(text) or 
                self._is_sub_heading_level2(text) or 
                self._is_sub_heading_level3(text))

    def _has_header1_style(self, paragraph):
        """ตรวจสอบว่าเป็น Header 1 style หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'heading 1' in style_name or 'header 1' in style_name:
                    return True
        except:
            pass
        return False

    def _has_header2_style(self, paragraph):
        """ตรวจสอบว่าเป็น Header 2 style หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'heading 2' in style_name or 'header 2' in style_name:
                    return True
        except:
            pass
        return False

    def _check_alignment(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบการจัดแนว - ไม่ตรวจสอบการจัดกึ่งกลางของหน้าปก"""
        text = paragraph.text.strip()
        if not text:
            return
        
        # ข้ามการตรวจสอบการจัดแนวสำหรับหน้าปก
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
        
        # หัวข้อหลัก หัวข้อใหญ่ และหัวข้อย่อยควรชิดซ้าย - แก้ไขให้ครอบคลุม
        if ((self._is_main_heading(text, paragraph) or self._is_major_heading(text, paragraph) or self._is_any_sub_heading(text)) 
            and self.standards['main_heading_left']):
            if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER or paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'การจัดแนวไม่ถูกต้อง',
                    'description': 'หัวข้อควรจัดแนวชิดซ้าย',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, "หัวข้อควรจัดแนวชิดซ้าย")

    def _check_indentation(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบการเยื้อง - แก้ไขให้ไม่ตรวจสอบการเยื้องของ caption"""
        text = paragraph.text.strip()
        if not text:
            return
        
        # ข้ามการตรวจสอบสำหรับหน้าปก
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
        
        # ข้ามการตรวจสอบสำหรับส่วนที่ไม่ต้องเยื้อง
        if self._should_skip_indent_check(text):
            return
        
        # ข้ามหัวข้อทุกประเภท - แก้ไขให้ครอบคลุม
        if (self._is_main_heading(text, paragraph) or self._is_major_heading(text, paragraph) or self._is_any_sub_heading(text)):
            return
        
        # ข้ามการอ้างอิงตารางและรูปภาพ (ในเนื้อหา ไม่ใช่ caption)
        if self._contains_table_reference(text) or self._contains_picture_reference(text):
            return
        
        # เนื้อหาทั่วไปควรมีการเยื้องย่อหน้า
        if self.standards['paragraph_indent']:
            if paragraph.paragraph_format.first_line_indent is None:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'การเยื้องไม่ถูกต้อง',
                    'description': 'เนื้อหาควรมีการเยื้องย่อหน้า (ประมาณ 0.5 นิ้ว)',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, "เนื้อหาควรมีการเยื้องย่อหน้า (กด Tab หรือตั้งค่า First Line Indent 0.5 นิ้ว)")

    def _check_font_sizes(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบขนาดฟอนต์ - แก้ไขให้จำแนกประเภทหัวข้อถูกต้อง"""
        text = paragraph.text.strip()
        if not text:
            return
        
        # ข้ามหน้าปก (จะตรวจสอบแยกใน _check_cover_page_detailed)
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
            
        expected_size = None
        section_type = None
        
        # 1. ชื่อบท - 18pt Bold
        if self._is_main_heading(text, paragraph) or self._has_header1_style(paragraph):
            expected_size = 18
            section_type = "ชื่อบท"
        # 2. หัวข้อใหญ่ (1. 2. 3.) - 16pt Bold
        elif self._is_major_heading(text, paragraph):
            expected_size = 16
            section_type = "หัวข้อใหญ่"
        # 3. หัวข้อย่อยทุกระดับ (1.1, 1.1.1, (1)) - 14pt ปกติ
        elif self._is_any_sub_heading(text):
            expected_size = 14
            section_type = "หัวข้อย่อย"
        # 4. เนื้อหาทั่วไป - 14pt ปกติ
        else:
            expected_size = 14
            section_type = "เนื้อหา"
        
        # ตรวจสอบขนาดฟอนต์ในทุก run
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                actual_size = None
                
                try:
                    if run.font.size and hasattr(run.font.size, 'pt'):
                        actual_size = run.font.size.pt
                    elif paragraph.style and paragraph.style.font and paragraph.style.font.size:
                        if hasattr(paragraph.style.font.size, 'pt'):
                            actual_size = paragraph.style.font.size.pt
                except AttributeError:
                    continue
                
                # ให้ความผิดพลาด ±1pt สำหรับความยืดหยุ่น
                if actual_size and abs(actual_size - expected_size) > 1:
                    self.errors.append({
                        'paragraph': para_num,
                        'type': 'ขนาดฟอนต์ไม่ถูกต้อง',
                        'description': f'{section_type}ควรใช้ขนาด {expected_size}pt แต่พบขนาด {actual_size}pt',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, f"{section_type}ควรใช้ขนาด {expected_size}pt (ปัจจุบัน {actual_size}pt)")
                    return

    def _check_font_bold(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบความหนาของฟอนต์ - แก้ไขให้จำแนกประเภทหัวข้อถูกต้อง"""
        text = paragraph.text.strip()
        if not text or self._should_skip_bold_check(text):
            return
        
        # ข้ามหน้าปก (ตรวจสอบแยกแล้ว)
        if self._is_cover_page_section(text, paragraph_index):
            return
        
        # ข้ามการตรวจสอบสำหรับ caption
        if self._is_caption_style(paragraph):
            return
        
        should_be_bold = None
        section_type = None
        
        # 1. ชื่อบท - ควรเป็น Bold
        if self._is_main_heading(text, paragraph) or self._has_header1_style(paragraph):
            should_be_bold = True
            section_type = "ชื่อบท"
        # 2. หัวข้อใหญ่ (1. 2. 3.) - ควรเป็น Bold
        elif self._is_major_heading(text, paragraph):
            should_be_bold = True
            section_type = "หัวข้อใหญ่"
        # 3. หัวข้อย่อยทุกระดับ - ควรเป็นตัวปกติ (ไม่หนา)
        elif self._is_any_sub_heading(text):
            should_be_bold = False
            section_type = "หัวข้อย่อย"
        
        # ตรวจสอบความหนาเฉพาะเมื่อมีการกำหนด
        if should_be_bold is not None and section_type:
            has_bold_text = self._is_paragraph_bold(paragraph) or self._has_bold_style(paragraph)
            
            if should_be_bold and not has_bold_text:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'ความหนาฟอนต์ไม่ถูกต้อง',
                    'description': f'{section_type}ควรเป็นตัวหนา (Bold)',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, f"{section_type}ควรเป็นตัวหนา (Bold)")
            elif not should_be_bold and has_bold_text:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'ความหนาฟอนต์ไม่ถูกต้อง',
                    'description': f'{section_type}ควรเป็นตัวปกติ (ไม่หนา)',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, f"{section_type}ควรเป็นตัวปกติ (ไม่หนา)")

    def _is_sub_heading(self, text):
        """ตรวจสอบว่าเป็นหัวข้อย่อยหรือไม่ - รวมทุกระดับ (เพื่อความเข้ากันได้กับโค้ดเดิม)"""
        return self._is_any_sub_heading(text)

    def _is_project_title_improved(self, text, index, context):
        """ตรวจสอบว่าเป็นชื่อโครงงานหรือไม่ - ปรับปรุงใหม่ให้ทำงานได้ดีขึ้น"""
        text_clean = text.strip()
        
        # Debug: แสดงข้อมูลการตรวจสอบ
        # print(f"  Checking if project title: '{text_clean[:30]}...' (length: {len(text_clean)})")
        
        # เงื่อนไขพื้นฐาน: ความยาวมากกว่า 8 ตัวอักษร
        if len(text_clean) < 8:
            # print(f"    -> Too short: {len(text_clean)} < 8")
            return False
        
        # คำสำคัญที่บ่งบอกว่าไม่ใช่ชื่อโครงงาน
        non_title_keywords = [
            'โดย', 'อาจารย์', 'ที่ปรึกษา', 'ภาค', 'สาขา', 'รายงาน', 'วิทยาลัย', 
            'มหาวิทยาลัย', 'เดือน', 'พ.ศ', 'ปีการศึกษา', 'การศึกษา', 'cs/it', 
            'cs ', 'เอกสาร', 'ฉบับสมบูรณ์', 'ความก้าวหน้า', 'โครงงานฉบับสมบูรณ์'
        ]
        
        for keyword in non_title_keywords:
            if keyword.lower() in text_clean.lower():
                # print(f"    -> Contains non-title keyword: {keyword}")
                return False
        
        # ไม่ใช่รหัสนักศึกษา
        if re.search(r'\d{9,10}-\d', text_clean):
            # print(f"    -> Contains student ID pattern")
            return False
        
        # ไม่เริ่มต้นด้วยตัวเลขหรือสัญลักษณ์พิเศษ (ยกเว้นวงเล็บ)
        if re.match(r'^[\d\W&&[^(]]', text_clean):
            # print(f"    -> Starts with number or special character")
            return False
        
        # ตรวจสอบตำแหน่งใน context
        if context.get('project_type_index', -1) >= 0 and context.get('author_section_start', -1) >= 0:
            # ถ้าอยู่ระหว่างประเภทโครงงานและส่วนผู้แต่ง
            if context['project_type_index'] < index < context['author_section_start']:
                # print(f"    -> Position check passed: {context['project_type_index']} < {index} < {context['author_section_start']}")
                return True
        
        # ตรวจสอบรูปแบบภาษาอังกฤษที่อาจเป็นชื่อโครงงาน
        if re.search(r'[A-Z][a-z]+.*[A-Z]', text_clean):  # มีการผสมตัวใหญ่-เล็กแบบชื่อเรื่อง
            # print(f"    -> English title pattern detected")
            return True
        
        # ตรวจสอบรูปแบบภาษาไทยที่อาจเป็นชื่อโครงงาน
        if re.search(r'[ก-๙].*ระบบ|[ก-๙].*การ|[ก-๙].*โครงงาน', text_clean):
            # print(f"    -> Thai title pattern detected")
            return True
        
        # เพิ่มการตรวจสอบความยาวที่เหมาะสม (ชื่อโครงงานมักจะยาว)
        if len(text_clean) > 15 and not any(char.isdigit() for char in text_clean[:5]):
            # print(f"    -> Long text without numbers at start")
            return True
        
        # print(f"    -> No conditions matched")
        return False

    def _debug_cover_structure(self, cover_paragraphs, cover_context):
        """ฟังก์ชัน debug สำหรับดูโครงสร้างหน้าปก"""
        print("=== Debug Cover Structure ===")
        for i, para in enumerate(cover_paragraphs):
            text = para.text.strip()
            if text:
                print(f"Index {i}: '{text[:50]}...' (length: {len(text)})")
                if i in cover_context['project_title_candidates']:
                    print(f"  -> ระบุเป็นชื่อโครงงาน candidate")
                if self._is_project_title_improved(text, i, cover_context):
                    print(f"  -> ผ่านการตรวจสอบ _is_project_title_improved")
        
        print(f"Project type index: {cover_context['project_type_index']}")
        print(f"Author section start: {cover_context['author_section_start']}")
        print(f"Project title candidates: {cover_context['project_title_candidates']}")
        print("=============================")

    def _get_expected_cover_font_size_with_context(self, text, index, context):
        """กำหนดขนาดฟอนต์หน้าปกโดยใช้ context - แก้ไขให้ตรงกับมาตรฐาน PDF"""
        text_clean = text.strip()
        text_lower = text_clean.lower()
        
        # Debug: แสดงข้อมูลการตรวจสอบ
        # print(f"Checking font size for index {index}: '{text_clean[:30]}...'")
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม = 26pt
        if re.search(r'cs\s*\d{4}', text_lower) or 'cs/it' in text_lower:
            # print(f"  -> CS code detected: 26pt")
            return 26
            
        # 2. ชื่อเล่ม = 20pt
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า']):
            # print(f"  -> Project type detected: 20pt")
            return 20
            
        # 3. ชื่อโครงงาน (ใช้ context) = 20pt - ปรับปรุงลำดับการตรวจสอบ
        elif (index in context['project_title_candidates'] or 
            self._is_project_title_improved(text_clean, index, context)):
            # print(f"  -> Project title detected: 20pt")
            return 20
            
        # 4. คำว่า "โดย" และชื่อผู้ทำ = 18pt
        elif ('โดย' in text_clean or re.search(r'\d{9,10}-\d', text_clean) or 
            (context['author_section_start'] > 0 and 
            context['author_section_start'] <= index < context.get('advisor_section_start', float('inf')))):
            # print(f"  -> Author section detected: 18pt")
            return 18
            
        # 5. อาจารย์ที่ปรึกษา = 18pt
        elif (any(keyword in text_clean for keyword in ['อาจารย์', 'ที่ปรึกษา']) or
            (context['advisor_section_start'] > 0 and 
            context['advisor_section_start'] <= index < context.get('bottom_section_start', float('inf')))):
            # print(f"  -> Advisor section detected: 18pt")
            return 18
            
        # 6. ส่วนด้านล่าง = 16pt
        elif (context['bottom_section_start'] > 0 and index >= context['bottom_section_start']) or \
            any(keyword in text_clean for keyword in ['สาขา', 'วิทยาลัย', 'มหาวิทยาลัย', 'ภาคเรียน', 'เดือน', 'พ.ศ', 'รายงานนี้', 'การศึกษา']):
            # print(f"  -> Bottom section detected: 16pt")
            return 16
        
        # print(f"  -> No specific rule matched")
        return None

    def _analyze_cover_structure(self, cover_paragraphs):
        """วิเคราะห์โครงสร้างหน้าปกเพื่อระบุตำแหน่งของแต่ละส่วน - ปรับปรุงให้แม่นยำขึ้น"""
        context = {
            'cs_code_index': -1,
            'project_type_index': -1,
            'author_section_start': -1,
            'advisor_section_start': -1,
            'bottom_section_start': -1,
            'project_title_candidates': []
        }
        
        for i, para in enumerate(cover_paragraphs):
            text = para.text.strip()
            text_lower = text.lower()
            
            # หารหัสสาขาและปี
            if re.search(r'cs\s*\d{4}', text_lower) or 'cs/it' in text_lower:
                context['cs_code_index'] = i
            
            # หาประเภทโครงงาน - ตรวจสอบให้แม่นยำขึ้น
            elif ('โครงงานฉบับสมบูรณ์' in text or 'เอกสารโครงงาน' in text or 
                'รายงานความก้าวหน้า' in text):
                context['project_type_index'] = i
            
            # หาส่วนผู้แต่ง - ปรับให้ยืดหยุ่นขึ้น
            elif ('โดย' in text and len(text.strip()) <= 10):  # คำว่า "โดย" อย่างเดียวหรือใกล้เคียง
                context['author_section_start'] = i
            
            # หาส่วนอาจารย์ที่ปรึกษา
            elif 'อาจารย์ที่ปรึกษา' in text:
                context['advisor_section_start'] = i
            
            # หาส่วนข้อมูลด้านล่าง
            elif any(keyword in text for keyword in ['รายงานนี้เป็นส่วนหนึ่ง', 'สาขาวิชา', 'วิทยาลัย', 'มหาวิทยาลัย']):
                if context['bottom_section_start'] == -1:
                    context['bottom_section_start'] = i
        
        # ระบุชื่อโครงงานจากตำแหน่ง - ปรับปรุงการค้นหาให้แม่นยำขึ้น
        if context['project_type_index'] >= 0 and context['author_section_start'] >= 0:
            project_start = context['project_type_index'] + 1
            project_end = context['author_section_start']
            
            for i in range(project_start, project_end):
                if i < len(cover_paragraphs):
                    text = cover_paragraphs[i].text.strip()
                    # ปรับเงื่อนไข: ให้ยืดหยุ่นขึ้นในการระบุชื่อโครงงาน
                    if (len(text) > 5 and  # ลดจาก 10 เป็น 5 เพื่อให้ครอบคลุมมากขึ้น
                        not any(keyword in text.lower() for keyword in ['โดย', 'อาจารย์', 'cs/it', 'cs ', 'รหัส']) and
                        not re.search(r'\d{9,10}-\d', text) and  # ไม่ใช่รหัสนักศึกษา
                        not text.isdigit() and  # ไม่ใช่ตัวเลขเพียงอย่างเดียว
                        text not in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน']):  # ไม่ใช่ชื่อประเภทเอกสาร
                        context['project_title_candidates'].append(i)
        
        return context
    
    def _check_cover_page_detailed(self, doc):
        """ตรวจสอบหน้าปกตามโครงสร้างใหม่ - เพิ่ม debug และปรับปรุงการระบุชื่อโครงงาน"""
        cover_end_index = self._find_cover_end_index(doc)
        cover_paragraphs = [p for p in doc.paragraphs[:cover_end_index] if p.text.strip()]
        
        # สร้าง context สำหรับการวิเคราะห์ตำแหน่ง
        cover_context = self._analyze_cover_structure(cover_paragraphs)
        
        # Debug: แสดงโครงสร้างหน้าปก (comment out ในการใช้งานจริง)
        # self._debug_cover_structure(cover_paragraphs, cover_context)
        
        # ตรวจสอบเฉพาะขนาดฟอนต์และความหนาของหน้าปก
        for i, para in enumerate(cover_paragraphs):
            text = para.text.strip()
            if not text:
                continue
            
            # ใช้ context ในการกำหนดขนาดฟอนต์
            expected_size = self._get_expected_cover_font_size_with_context(text, i, cover_context)
            if expected_size:
                actual_size = self._get_paragraph_font_size(para)
                if actual_size and abs(actual_size - expected_size) > 2:  # ให้ความผิดพลาด ±2pt
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'ขนาดฟอนต์หน้าปกไม่ถูกต้อง',
                        'description': f'ข้อความ "{text[:30]}..." ในหน้าปกควรใช้ขนาด {expected_size}pt แต่พบ {actual_size}pt',
                        'paragraph_obj': para
                    })
                    self._highlight_paragraph(para)
                    self._add_comment_to_paragraph(para, f'ขนาดฟอนต์ควร {expected_size}pt')
            
            # ตรวจสอบความหนาเฉพาะส่วนที่ควรเป็น Bold
            expected_bold = self._should_cover_text_be_bold_with_context(text, i, cover_context)
            if expected_bold is not None:
                is_bold = self._is_paragraph_bold(para) or self._has_bold_style(para)
                if expected_bold and not is_bold:
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'ความหนาฟอนต์หน้าปกไม่ถูกต้อง',
                        'description': f'ข้อความ "{text[:30]}..." ในหน้าปกควรเป็นตัวหนา',
                        'paragraph_obj': para
                    })
                    self._highlight_paragraph(para)
                    self._add_comment_to_paragraph(para, 'ข้อความนี้ในหน้าปกควรเป็นตัวหนา')
                elif expected_bold == False and is_bold:
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'ความหนาฟอนต์หน้าปกไม่ถูกต้อง',
                        'description': f'ข้อความ "{text[:30]}..." ในหน้าปกควรเป็นตัวปกติ (ไม่หนา)',
                        'paragraph_obj': para
                    })
                    self._highlight_paragraph(para)
                    self._add_comment_to_paragraph(para, 'ข้อความนี้ในหน้าปกควรเป็นตัวปกติ (ไม่หนา)')

    def _should_cover_text_be_bold_with_context(self, text, index, context):
        """กำหนดว่าข้อความในหน้าปกควรเป็น Bold หรือไม่ โดยใช้ context"""
        text_clean = text.strip()
        text_lower = text_clean.lower()
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม = หนา
        if re.search(r'cs\s*\d{4}', text_lower) or 'cs/it' in text_lower:
            return True
            
        # 2. ชื่อเล่ม = หนา
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า']):
            return True
            
        # 3. ชื่อโครงงาน = ปกติ (ไม่หนา) - ปรับปรุงลำดับการตรวจสอบ
        elif (index in context['project_title_candidates'] or 
            self._is_project_title_improved(text_clean, index, context)):
            return False
            
        # 4. ส่วนอื่นๆ = ปกติ (ไม่หนา)
        else:
            return False
    
    def _get_expected_cover_font_size_from_pdf(self, text):
        """กำหนดขนาดฟอนต์หน้าปกตามโครงสร้างที่อธิบาย - ปรับปรุงใหม่"""
        text_clean = text.strip()   
        text_lower = text_clean.lower()
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม เช่น CS 2567/CS-59 = 26pt หนา
        if re.search(r'cs\s*\d{4}', text_lower) or re.search(r'cs/it.*\d{4}', text_lower) or 'cs/it' in text_lower:
            return 26
            
        # 2. ชื่อเล่ม เช่น โครงงานฉบับสมบูรณ์, รายงานความก้าวหน้า = 20pt หนา  
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า', 'โครงงาน']):
            return 20
            
        # 3. ชื่อเรื่องโครงงานภาษาไทยและอังกฤษ = 20pt ปกติ
        # ปรับปรุง: ใช้วิธีการระบุตำแหน่งที่แม่นยำขึ้น
        elif self._is_project_title(text_clean):
            return 20
            
        # 4. คำว่า "โดย" และชื่อผู้ทำโครงงาน = 18pt ปกติ
        elif 'โดย' in text_clean or re.search(r'\d{9,10}-\d', text_clean):
            return 18
            
        # 5. ชื่ออาจารย์ที่ปรึกษา = 18pt ปกติ  
        elif any(keyword in text_clean for keyword in ['อาจารย์', 'ที่ปรึกษา']):
            return 18
            
        # 6. ส่วนด้านล่าง (ข้อมูลสาขา, ภาคเรียน, เดือน, ปี) = 16pt ปกติ
        elif any(keyword in text_clean for keyword in ['สาขา', 'วิทยาลัย', 'มหาวิทยาลัย', 'ภาคเรียน', 'เดือน', 'พ.ศ', 'รายงานนี้', 'การศึกษา', 'ปีการศึกษา']):
            return 16
        
        return None
        
    def _should_cover_text_be_bold_from_pdf(self, text):
        """กำหนดว่าข้อความในหน้าปกควรเป็น Bold หรือไม่ - ปรับปรุงใหม่"""
        text_clean = text.strip()
        text_lower = text_clean.lower()
        
        # 1. รหัสสาขา พ.ศ./กลุ่ม เช่น CS 2567/CS-59 = หนา
        if re.search(r'cs\s*\d{4}', text_lower) or re.search(r'cs/it.*\d{4}', text_lower) or 'cs/it' in text_lower:
            return True
            
        # 2. ชื่อเล่ม เช่น โครงงานฉบับสมบูรณ์ = หนา
        elif any(keyword in text_clean for keyword in ['โครงงานฉบับสมบูรณ์', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า']):
            return True
            
        # 3. ชื่อเรื่องโครงงาน = ปกติ (ไม่หนา)
        elif self._is_project_title(text_clean):
            return False
            
        # 4. ชื่อผู้ทำโครงงาน = ปกติ
        elif 'โดย' in text_clean or re.search(r'\d{9,10}-\d', text_clean):
            return False
            
        # 5. ชื่ออาจารย์ที่ปรึกษา = ปกติ
        elif any(keyword in text_clean for keyword in ['อาจารย์', 'ที่ปรึกษา']):
            return False
            
        # 6. ส่วนด้านล่าง = ปกติ
        elif any(keyword in text_clean for keyword in ['สาขา', 'วิทยาลัย', 'มหาวิทยาลัย', 'ภาคเรียน', 'เดือน', 'พ.ศ', 'รายงานนี้', 'การศึกษา', 'ปีการศึกษา']):
            return False
        
        return None
    
    def _is_project_title(self, text):
        """ตรวจสอบว่าเป็นชื่อโครงงานหรือไม่"""
        text_clean = text.strip()
        
        # เงื่อนไขสำหรับระบุชื่อโครงงาน:
        # 1. ความยาวมากกว่า 15 ตัวอักษร (ชื่อโครงงานมักจะยาว)
        # 2. ไม่มีคำสำคัญที่บ่งบอกว่าไม่ใช่ชื่อโครงงาน
        # 3. ไม่ใช่รหัสนักศึกษา
        # 4. ไม่ใช่ข้อมูลส่วนอื่นๆ
        
        if len(text_clean) < 15:
            return False
        
        # คำสำคัญที่บ่งบอกว่าไม่ใช่ชื่อโครงงาน
        non_title_keywords = [
            'โดย', 'อาจารย์', 'ที่ปรึกษา', 'ภาค', 'สาขา', 'รายงาน', 'วิทยาลัย', 
            'มหาวิทยาลัย', 'เดือน', 'พ.ศ', 'ปีการศึกษา', 'การศึกษา', 'cs/it', 
            'cs ', 'โครงงาน', 'เอกสาร', 'ฉบับสมบูรณ์', 'ความก้าวหน้า'
        ]
        
        for keyword in non_title_keywords:
            if keyword.lower() in text_clean.lower():
                return False
        
        # ไม่ใช่รหัสนักศึกษา
        if re.search(r'\d{9,10}-\d', text_clean):
            return False
        
        # ไม่เริ่มต้นด้วยตัวเลขหรือสัญลักษณ์พิเศษ
        if re.match(r'^[\d\W]', text_clean):
            return False
        
        return True

    def _find_cover_end_index(self, doc):
        """หาจุดสิ้นสุดของหน้าปก - ปรับปรุงให้รองรับหน้าปก 2 หน้าที่ไม่มีเลขหน้า"""
        cover_end_index = 0
        found_page_number = False
        
        for i, para in enumerate(doc.paragraphs[:60]):  # เพิ่มขอบเขตการค้นหา
            text = para.text.strip().lower()
            
            # ตรวจสอบเลขหน้า - หากเจอเลขหน้าแสดงว่าจบหน้าปกแล้ว
            if re.search(r'^\s*\d+\s*$', para.text.strip()) and i > 10:  # เลขหน้าที่ปรากฏหลังจากเนื้อหาหน้าปกบ้าง
                found_page_number = True
                cover_end_index = i
                break
            
            # คำสำคัญที่บ่งชี้ว่าจบหน้าปกแล้ว
            end_cover_keywords = [
                'บทคัดย่อ', 'abstract', 'สารบัญ', 'บทที่', 'กิตติกรรม', 
                'acknowledgement', 'table of contents', 'contents', 
                'ประกาศนียบัตร', 'certificate', 'คำนำ', 'preface'
            ]
            
            if any(keyword in text for keyword in end_cover_keywords):
                cover_end_index = i
                break
        
        # หากไม่เจอคำสำคัญใดๆ ให้ใช้ค่า default สำหรับหน้าปก 2 หน้า
        if cover_end_index == 0:
            cover_end_index = 40  # ประมาณ 2 หน้า
        
        return cover_end_index

    def _check_references(self, doc):
        """ตรวจสอบการอ้างอิงตารางและรูปภาพ - แก้ไขให้ตรวจสอบจาก caption style"""
        for i, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            
            # ตรวจสอบว่าเป็น caption หรือไม่จาก style
            is_caption = self._is_caption_style(paragraph)
            
            if is_caption:
                # ถ้าเป็น caption แล้วข้ามการตรวจสอบการเยื้อง
                continue
            
            # ตรวจสอบการอ้างอิงตารางในเนื้อหา (ไม่ใช่ caption)
            if self._contains_table_reference(text) and not is_caption:
                if not self._is_valid_table_reference_format(text):
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'รูปแบบการอ้างอิงตารางไม่ถูกต้อง',
                        'description': 'การอ้างอิงตารางควรเป็น "ตารางที่ X" หรือ "ตารางที่ X.X"',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, 'การอ้างอิงตารางควรเป็น "ตารางที่ X" หรือ "ตารางที่ X.X"')
            
            # ตรวจสอบการอ้างอิงรูปภาพในเนื้อหา (ไม่ใช่ caption)
            if self._contains_picture_reference(text) and not is_caption:
                if not self._is_valid_picture_reference_format(text):
                    self.errors.append({
                        'paragraph': i+1,
                        'type': 'รูปแบบการอ้างอิงรูปภาพไม่ถูกต้อง',
                        'description': 'การอ้างอิงรูปภาพควรเป็น "ภาพที่ X" หรือ "ภาพที่ X.X"',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, 'การอ้างอิงรูปภาพควรเป็น "ภาพที่ X" หรือ "ภาพที่ X.X"')

    def _is_caption_style(self, paragraph):
        """ตรวจสอบว่าเป็น caption style หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'caption' in style_name:
                    return True
        except:
            pass
        return False

    def _contains_table_reference(self, text):
        """ตรวจสอบว่าข้อความมีการอ้างอิงตารางหรือไม่"""
        patterns = [
            r'ตารางที่\s*\d+',
            r'ตารางที่\s*\d+\.\d+',
            r'table\s*\d+',
            r'Table\s*\d+',
        ]
        for pattern in patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return True
        return False
    
    def _contains_picture_reference(self, text):
        """ตรวจสอบว่าข้อความมีการอ้างอิงรูปภาพหรือไม่"""
        patterns = [
            r'ภาพที่\s*\d+',
            r'ภาพที่\s*\d+\.\d+',
            r'รูปที่\s*\d+',
            r'รูปที่\s*\d+\.\d+',
            r'figure\s*\d+',
            r'Figure\s*\d+',
        ]
        for pattern in patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return True
        return False
    
    def _is_valid_table_reference_format(self, text):
        """ตรวจสอบรูปแบบการอ้างอิงตารางว่าถูกต้องหรือไม่"""
        valid_patterns = [
            r'^ตารางที่\s*\d+\s',
            r'^ตารางที่\s*\d+\.\d+\s',
            r'^ตารางที่\s*\d+',
            r'^ตารางที่\s*\d+\.\d+',
        ]
        for pattern in valid_patterns:
            if re.search(pattern, text):
                return True
        return False
    
    def _is_valid_picture_reference_format(self, text):
        """ตรวจสอบรูปแบบการอ้างอิงรูปภาพว่าถูกต้องหรือไม่"""
        valid_patterns = [
            r'^ภาพที่\s*\d+\s',
            r'^ภาพที่\s*\d+\.\d+\s',
            r'^ภาพที่\s*\d+',
            r'^ภาพที่\s*\d+\.\d+',
        ]
        for pattern in valid_patterns:
            if re.search(pattern, text):
                return True
        return False

    def check_document_object(self, doc):
        """ตรวจสอบ Document object ที่ผ่านเข้ามา"""
        self.errors = []
        self.tables = []
        self.pictures = []
        
        # ตรวจหาตารางและรูปภาพ
        self._extract_tables_and_pictures(doc)
        
        # ตรวจสอบหน้าปก
        self._check_cover_page_detailed(doc)
        
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip():
                # ข้ามหน้าปกในการตรวจสอบปกติ
                if not self._is_cover_page_section(paragraph.text, i):
                    self._check_paragraph(paragraph, i + 1, i)
        
        # ตรวจสอบการอ้างอิงตารางและรูปภาพ
        self._check_references(doc)
        
        return self.errors
    
    def _get_paragraph_font_size(self, paragraph):
        """ดึงขนาดฟอนต์จากย่อหน้า"""
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                if run.font.size and hasattr(run.font.size, 'pt'):
                    return run.font.size.pt
        
        try:
            if paragraph.style and paragraph.style.font and paragraph.style.font.size:
                if hasattr(paragraph.style.font.size, 'pt'):
                    return paragraph.style.font.size.pt
        except:
            pass
            
        return None
    
    def _is_paragraph_bold(self, paragraph):
        """ตรวจสอบว่าย่อหน้าเป็นตัวหนาหรือไม่"""
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                if run.font.bold:
                    return True
        return False

    def _has_bold_style(self, paragraph):
        """ตรวจสอบว่าย่อหน้ามี style ที่เป็น bold หรือไม่"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'font'):
                if paragraph.style.font.bold:
                    return True
        except:
            pass
        return False
    
    def _check_paragraph(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบย่อหน้าแต่ละย่อหน้า"""
        self._check_thai_font(paragraph, para_num, paragraph_index)
        self._check_font_sizes(paragraph, para_num, paragraph_index)
        self._check_font_bold(paragraph, para_num, paragraph_index)
        self._check_alignment(paragraph, para_num, paragraph_index)
        self._check_indentation(paragraph, para_num, paragraph_index)
        self._check_spacing(paragraph, para_num, paragraph_index)
        self._check_heading_numbering(paragraph, para_num, paragraph_index)
    
    def _is_cover_page_section(self, text, paragraph_index):
        """ตรวจสอบว่าเป็นส่วนของหน้าปกหรือไม่ - แก้ไขให้รองรับหน้าปก 2 หน้า"""
        # เพิ่มขอบเขตการตรวจสอบหน้าปกให้มากขึ้น (รองรับ 2 หน้า)
        if paragraph_index > 50:  # เพิ่มจาก 25 เป็น 50
            return False
        
        # ตรวจสอบคำสำคัญที่บ่งบอกว่าไม่ใช่หน้าปก
        non_cover_keywords = ['บทคัดย่อ', 'abstract', 'สารบัญ', 'บทที่', 'กิตติกรรม', 'acknowledgement', 'table of contents', 'contents', 'ประกาศนียบัตร', 'certificate']
        for keyword in non_cover_keywords:
            if keyword.lower() in text.lower():
                return False
        
        # ตรวจสอบคำสำคัญที่บ่งบอกว่าเป็นหน้าปก
        cover_keywords_from_standards = self.standards.get('cover_sections', [])
        for keyword in cover_keywords_from_standards:
            if keyword in text:
                return True
        
        # ตรวจสอบรหัสนักศึกษา (รูปแบบ ตัวเลข 9-10 หลัก ตามด้วย - และตัวเลข 1 หลัก)
        if re.search(r'\d{9,10}-\d', text):
            return True
        
        # ตรวจสอบคำสำคัญเพิ่มเติมสำหรับหน้าปก
        additional_cover_keywords = [
            'cs/it', 'cs ', 'ภาคเรียนที่', 'ปีการศึกษา', 'เดือน', 'พ.ศ.', 'ครั้งที่', 
            'วิทยาลัยการคอมพิวเตอร์', 'วิทยาลัยคอมพิวเตอร์', 'มหาวิทยาลัยขอนแก่น', 
            'โครงงาน', 'เอกสารโครงงาน', 'รายงานความก้าวหน้า', 'โดย', 'อาจารย์ที่ปรึกษา',
            'รายงานนี้เป็นส่วนหนึ่ง', 'การศึกษาวิชา'
        ]
        
        for keyword in additional_cover_keywords:
            if keyword.lower() in text.lower():
                return True
                
        return False
    
    def _has_header_style(self, paragraph):
        """ตรวจสอบว่าย่อหน้ามี Header style หรือไม่ - แก้ไขให้ตรวจสอบ Header 1 เป็นชื่อบท"""
        try:
            if paragraph.style and hasattr(paragraph.style, 'name'):
                style_name = paragraph.style.name.lower()
                if 'heading' in style_name or 'header' in style_name:
                    return True
        except:
            pass
        return False
    
    def _check_spacing(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบระยะห่างบรรทัด"""
        if not paragraph.text.strip():
            return
            
        line_spacing = paragraph.paragraph_format.line_spacing
        if line_spacing and line_spacing != 1.0:
            if isinstance(line_spacing, (int, float)) and line_spacing > 1.1:
                self.errors.append({
                    'paragraph': para_num,
                    'type': 'ระยะห่างบรรทัดไม่ถูกต้อง',
                    'description': f'ควรใช้ระยะห่างบรรทัด 1 เท่า (Single) แต่พบ {line_spacing}',
                    'paragraph_obj': paragraph
                })
                self._highlight_paragraph(paragraph)
                self._add_comment_to_paragraph(paragraph, f"ควรใช้ระยะห่างบรรทัด 1 เท่า (Single) - ปัจจุบัน {line_spacing}")
    
    def _check_heading_numbering(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบการใช้เลขหัวข้อ"""
        text = paragraph.text.strip()
        
        if re.match(r'^\d+(\.\d+)*\s*$', text):
            self.errors.append({
                'paragraph': para_num,
                'type': 'รูปแบบหัวข้อไม่ถูกต้อง',
                'description': 'หัวข้อควรมีเนื้อหาตามหลังเลขหัวข้อ',
                'paragraph_obj': paragraph
            })
            self._highlight_paragraph(paragraph)
            self._add_comment_to_paragraph(paragraph, "หัวข้อควรมีเนื้อหาตามหลังเลขหัวข้อ")
    
    def _should_skip_indent_check(self, text):
        """ตรวจสอบว่าควรข้ามการตรวจสอบการเยื้องหรือไม่"""
        for section in self.standards['no_indent_sections']:
            if section in text:
                return True
        return False
    
    def _should_skip_font_check(self, text):
        """ตรวจสอบว่าควรข้ามการตรวจสอบฟอนต์หรือไม่"""
        for section in self.standards['skip_font_check']:
            if section in text:
                return True
        return False
    
    def _should_skip_bold_check(self, text):
        """ตรวจสอบว่าควรข้ามการตรวจสอบความหนาหรือไม่"""
        for section in self.standards['skip_bold_check']:
            if section in text:
                return True
        return False
    
    def _add_comment_to_paragraph(self, paragraph, comment_text):
        """เพิ่ม comment ที่ท้ายย่อหน้า"""
        if "[❌ ข้อผิดพลาด:" in paragraph.text:
            return
        
        new_run = paragraph.add_run(f" [❌ ข้อผิดพลาด: {comment_text}]")
        new_run.font.color.rgb = RGBColor(255, 0, 0)
        new_run.font.bold = True
        new_run.font.size = Pt(12)
    
    def _highlight_paragraph(self, paragraph):
        """ทำ highlight ย่อหน้า"""
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                run.font.highlight_color = 7
    
    def _check_thai_font(self, paragraph, para_num, paragraph_index):
        """ตรวจสอบฟอนต์ภาษาไทย"""
        text = paragraph.text.strip()
        if not text or self._should_skip_font_check(text):
            return
        
        has_thai = any(ord(char) >= 0x0E00 and ord(char) <= 0x0E7F for char in text)
        if not has_thai:
            return
        
        for run in paragraph.runs:
            if run.text.strip() and "[❌ ข้อผิดพลาด:" not in run.text:
                font_name = run.font.name
                if font_name and font_name != self.standards['theme_font']:
                    self.errors.append({
                        'paragraph': para_num,
                        'type': 'ฟอนต์ไม่ถูกต้อง',
                        'description': f'ข้อความภาษาไทยควรใช้ฟอนต์ {self.standards["theme_font"]} (พบ {font_name})',
                        'paragraph_obj': paragraph
                    })
                    self._highlight_paragraph(paragraph)
                    self._add_comment_to_paragraph(paragraph, f"ข้อความภาษาไทยควรใช้ฟอนต์ {self.standards['theme_font']} (ปัจจุบันใช้ {font_name})")
                    return
    
    def get_error_statistics(self):
        """สร้างสถิติข้อผิดพลาด"""
        error_types = {}
        for error in self.errors:
            error_type = error['type']
            error_types[error_type] = error_types.get(error_type, 0) + 1
        return error_types