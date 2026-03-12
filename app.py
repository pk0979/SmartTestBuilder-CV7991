import streamlit as st
import pandas as pd
import random
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- HÀM TẠO BẢNG MA TRẬN ĐẶC TẢ ---
def create_matrix_table(doc, matrix_data):
    doc.add_paragraph("BẢNG MA TRẬN ĐẶC TẢ ĐỀ KIỂM TRA").runs[0].bold = True
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    
    # Header bảng
    hdr_cells = table.rows[0].cells
    headers = ['STT', 'Chủ đề/Chương', 'Nhận biết', 'Thông hiểu', 'Vận dụng', 'Vận dụng cao']
    for i, name in enumerate(headers):
        hdr_cells[i].text = name
        hdr_cells[i].paragraphs[0].runs[0].bold = True

    # Đổ dữ liệu từ matrix_data (dict gom nhóm theo chương)
    for idx, (chapter, levels) in enumerate(matrix_data.items()):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx + 1)
        row_cells[1].text = chapter
        row_cells[2].text = str(levels.get('NB', 0))
        row_cells[3].text = str(levels.get('TH', 0))
        row_cells[4].text = str(levels.get('VD', 0))
        row_cells[5].text = str(levels.get('VDC', 0))

# --- CẬP NHẬT HÀM XUẤT WORD CHÍNH ---
def export_full_exam(exam_list, info, matrix_summary):
    doc = Document()
    
    # 1. Header trường lớp
    p = doc.add_paragraph()
    p.add_run(f"{info['school'].upper()}\n").bold = True
    p.add_run(f"GV: {info['teacher']}").italic = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # 2. Chèn Ma trận đặc tả
    create_matrix_table(doc, matrix_summary)
    doc.add_page_break()
    
    # 3. Nội dung đề thi (Giữ nguyên logic cũ của bạn)
    doc.add_paragraph(f"ĐỀ KIỂM TRA MÔN {info['subject'].upper()}").runs[0].bold = True
    # ... (Phần code trộn đề và in câu hỏi như ở lượt trước) ...
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- GIAO DIỆN STREAMLIT ---
# (Thêm phần thống kê matrix_summary từ các câu đã chọn để truyền vào hàm trên)