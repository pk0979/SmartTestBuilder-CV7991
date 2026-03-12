import streamlit as st
import pandas as pd
import random
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Cấu hình trang
st.set_page_config(page_title="EduTest - CV7991", layout="wide")

def style_text(paragraph, text, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    return run

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

    # Đổ dữ liệu
    for idx, (chapter, levels) in enumerate(matrix_data.items()):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx + 1)
        row_cells[1].text = chapter
        row_cells[2].text = str(levels.get('NB', 0))
        row_cells[3].text = str(levels.get('TH', 0))
        row_cells[4].text = str(levels.get('VD', 0))
        row_cells[5].text = str(levels.get('VDC', 0))

# --- HÀM TRỘN ĐỀ VÀ PHƯƠNG ÁN ---
def shuffle_question(row):
    options = [('A', str(row['A'])), ('B', str(row['B'])), ('C', str(row['C'])), ('D', str(row['D']))]
    correct_content = str(row[row['Dap_an_dung']])
    random.shuffle(options)
    
    new_options = {}
    new_correct_label = ""
    labels = ['A', 'B', 'C', 'D']
    
    for i in range(4):
        new_label = labels[i]
        content = options[i][1]
        new_options[new_label] = content
        if content == correct_content:
            new_correct_label = new_label
            
    return {
        'Noi_dung': row['Noi_dung'],
        'Options': new_options,
        'Correct': new_correct_label,
        'Muc_do': row['Muc_do'],
        'Chuong': row.get('Chuong', 'Chung') # Lấy tên chương để làm ma trận
    }

# --- HÀM XUẤT FILE WORD CHUẨN ---
def export_full_exam(exam_list, info, matrix_summary):
    doc = Document()
    
    # Header
    header_table = doc.add_table(rows=1, cols=2)
    header_table.width = Inches(6)
    left_cell = header_table.cell(0, 0).paragraphs[0]
    style_text(left_cell, f"{info['school'].upper()}\n", bold=True)
    style_text(left_cell, f"GV: {info['teacher']}", italic=True)
    
    right_cell = header_table.cell(0, 1).paragraphs[0]
    right_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_text(right_cell, f"ĐỀ KIỂM TRA MÔN {info['subject'].upper()}\n", bold=True)
    style_text(right_cell, f"Thời gian làm bài: {info['time']} phút")

    doc.add_paragraph("_" * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Bảng Ma trận
    create_matrix_table(doc, matrix_summary)
    doc.add_page_break()
    
    # Nội dung đề thi
    doc.add_paragraph(f"ĐỀ KIỂM TRA MÔN {info['subject'].upper()}").runs[0].bold = True
    doc.add_paragraph("\nPHẦN I. TRẮC NGHIỆM").runs[0].bold = True
    
    for i, q in enumerate(exam_list):
        p = doc.add_paragraph()
        style_text(p, f"Câu {i+1}: ", bold=True)
        style_text(p, q['Noi_dung'])
        for label in ['A', 'B', 'C', 'D']:
            opt_p = doc.add_paragraph()
            opt_p.paragraph_format.left_indent = Inches(0.3)
            style_text(opt_p, f"{label}. ", bold=True)
            style_text(opt_p, q['Options'][label])

    # Bảng đáp án
    doc.add_page_break()
    doc.add_paragraph("BẢNG ĐÁP ÁN").runs[0].bold = True
    ans_table = doc.add_table(rows=2, cols=len(exam_list))
    ans_table.style = 'Table Grid'
    
    for i, q in enumerate(exam_list):
        ans_table.cell(0, i).text = str(i+1)
        ans_table.cell(1, i).text = q['Correct']

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- GIAO DIỆN CHÍNH ---
st.title("🎓 Trình Tạo Đề Kiểm Tra Tự Động (Chuẩn CV 7991)")
col1, col2 = st.columns([1, 2])

with col1:
    st.header("1. Thiết lập chung")
    school = st.text_input("Tên trường", "THCS Đa Phước")
    teacher = st.text_input("Tên giáo viên", "Tống Phước Khải")
    subject = st.text_input("Môn học", "Toán")
    duration = st.number_input("Thời gian (phút)", value=90)
    
    st.divider()
    uploaded_file = st.file_uploader("Tải lên Ngân hàng câu hỏi (Excel)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"Đã nhận {len(df)} câu hỏi từ file!")

    with col2:
        st.header("2. Cấu hình Ma trận đề")
        stats = df['Muc_do'].value_counts()
        
        c1, c2, c3, c4 = st.columns(4)
        n_nb = c1.number_input(f"Nhận biết (Có {stats.get('NB',0)})", 0, stats.get('NB',0), min(10, stats.get('NB',0)))
        n_th = c2.number_input(f"Thông hiểu (Có {stats.get('TH',0)})", 0, stats.get('TH',0), min(7, stats.get('TH',0)))
        n_vd = c3.number_input(f"Vận dụng (Có {stats.get('VD',0)})", 0, stats.get('VD',0), min(2, stats.get('VD',0)))
        n_vdc = c4.number_input(f"Vận dụng cao (Có {stats.get('VDC',0)})", 0, stats.get('VDC',0), min(1, stats.get('VDC',0)))

        if st.button("🚀 TẠO ĐỀ VÀ XUẤT FILE"):
            selected_rows = []
            for level, count in zip(['NB', 'TH', 'VD', 'VDC'], [n_nb, n_th, n_vd, n_vdc]):
                if count > 0:
                    pool = df[df['Muc_do'] == level]
                    selected_rows.append(pool.sample(n=count))
            
            if selected_rows:
                final_df = pd.concat(selected_rows)
                exam_list = [shuffle_question(row) for _, row in final_df.iterrows()]
                random.shuffle(exam_list)
                
                # Thống kê số lượng câu cho từng chương để vẽ Ma trận
                matrix_summary = {}
                for q in exam_list:
                    chuong = str(q['Chuong'])
                    muc_do = q['Muc_do']
                    if chuong not in matrix_summary:
                        matrix_summary[chuong] = {'NB':0, 'TH':0, 'VD':0, 'VDC':0}
                    matrix_summary[chuong][muc_do] += 1
                
                # Xuất Word
                info = {'school': school, 'teacher': teacher, 'subject': subject, 'time': duration}
                docx_file = export_full_exam(exam_list, info, matrix_summary)
                
                st.download_button(
                    label="📥 Tải Đề Kiểm Tra & Ma Trận (.docx)",
                    data=docx_file,
                    file_name=f"De_Kiem_Tra_{subject}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.balloons()
