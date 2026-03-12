import streamlit as st
import pandas as pd
import random
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="EduTest Pro - CV7991", page_icon="🎓", layout="wide")

# Custom CSS để giao diện đẹp hơn
st.markdown("""
<style>
    .stMetric { background-color: #f0f2f6; padding: 10px; border-radius: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 4px 4px 0 0; padding-top: 10px; padding-bottom: 10px; }
    .stTabs [aria-selected="true"] { background-color: #e6f0ff; border-bottom: 3px solid #0066cc; font-weight: bold;}
</style>
""", unsafe_allow_html=True)

# --- KHỞI TẠO SESSION STATE ---
if 'matrix_df' not in st.session_state:
    st.session_state.matrix_df = pd.DataFrame()

# --- HÀM TIỆN ÍCH ---
def style_text(paragraph, text, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    return run

# --- HÀM TRÍ TUỆ NHÂN TẠO (AI) PHÂN BỔ MA TRẬN ---
def ai_generate_matrix(df, total_q, strategy_ratio):
    chapters = df['Chuong'].unique()
    levels = ['NB', 'TH', 'VD', 'VDC']
    
    # Tính tổng target theo chiến lược
    targets = {lvl: int(total_q * r / 100) for lvl, r in zip(levels, strategy_ratio)}
    # Bù trừ sai số làm tròn cho VDC
    targets['VDC'] = total_q - sum([targets[l] for l in ['NB', 'TH', 'VD']])
    
    # Tạo bảng kế hoạch rỗng
    plan = {ch: {lvl: 0 for lvl in levels} for ch in chapters}
    
    # Lấy số lượng câu hỏi tối đa hiện có trong DB
    available = {ch: {lvl: len(df[(df['Chuong'] == ch) & (df['Muc_do'] == lvl)]) for lvl in levels} for ch in chapters}
    
    # Thuật toán phân bổ đều
    for lvl in levels:
        t = targets[lvl]
        while t > 0:
            # Kiểm tra xem còn câu hỏi nào ở mức độ này không
            total_avail_lvl = sum([available[ch][lvl] for ch in chapters])
            if total_avail_lvl == 0:
                st.warning(f"⚠️ Ngân hàng thiếu câu hỏi mức {lvl}! AI chỉ có thể xếp được {targets[lvl] - t}/{targets[lvl]} câu.")
                break
                
            for ch in chapters:
                if t > 0 and available[ch][lvl] > 0:
                    plan[ch][lvl] += 1
                    available[ch][lvl] -= 1
                    t -= 1
                    
    # Chuyển đổi thành DataFrame để hiển thị trên UI
    matrix_df = pd.DataFrame.from_dict(plan, orient='index')
    return matrix_df

# --- HÀM TRỘN CÂU HỎI ---
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
        'Chuong': row.get('Chuong', 'Chung')
    }

# --- HÀM XUẤT WORD ---
def export_full_exam(exam_list, info, matrix_df):
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
    
    # Bảng Ma trận đặc tả
    doc.add_paragraph("BẢNG MA TRẬN ĐẶC TẢ ĐỀ KIỂM TRA").runs[0].bold = True
    m_table = doc.add_table(rows=1, cols=6)
    m_table.style = 'Table Grid'
    hdr_cells = m_table.rows[0].cells
    for i, name in enumerate(['STT', 'Chủ đề/Chương', 'Nhận biết', 'Thông hiểu', 'Vận dụng', 'Vận dụng cao']):
        hdr_cells[i].text = name
        hdr_cells[i].paragraphs[0].runs[0].bold = True

    idx = 1
    for chapter, row in matrix_df.iterrows():
        row_cells = m_table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(chapter)
        row_cells[2].text = str(row.get('NB', 0))
        row_cells[3].text = str(row.get('TH', 0))
        row_cells[4].text = str(row.get('VD', 0))
        row_cells[5].text = str(row.get('VDC', 0))
        idx += 1
        
    doc.add_page_break()
    
    # Nội dung đề
    doc.add_paragraph(f"ĐỀ KIỂM TRA MÔN {info['subject'].upper()}").runs[0].bold = True
    doc.add_paragraph("\nPHẦN I. TRẮC NGHIỆM").runs[0].bold = True
    
    for i, q in enumerate(exam_list):
        p = doc.add_paragraph()
        style_text(p, f"Câu {i+1} ({q['Muc_do']}): ", bold=True)
        style_text(p, q['Noi_dung'])
        for label in ['A', 'B', 'C', 'D']:
            opt_p = doc.add_paragraph()
            opt_p.paragraph_format.left_indent = Inches(0.3)
            style_text(opt_p, f"{label}. ", bold=True)
            style_text(opt_p, q['Options'][label])

    # Đáp án
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

# --- UI CHÍNH ---
st.title("🎓 Hệ Thống Tạo Đề Thông Minh (Chuẩn CV 7991)")

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3074/3074058.png", width=100)
    st.header("⚙️ Thông tin chung")
    school = st.text_input("Tên trường", "THCS Đa Phước")
    teacher = st.text_input("Tên giáo viên", "Tống Phước Khải")
    subject = st.text_input("Môn học", "Toán")
    duration = st.number_input("Thời gian (phút)", value=90)
    st.info("Hệ thống tích hợp AI tự động phân bổ ma trận đặc tả theo chuẩn BGDĐT.")

# Quy trình 3 bước chuyên nghiệp
tab1, tab2, tab3 = st.tabs(["📁 1. Dữ liệu Ngân hàng", "🤖 2. Trợ lý AI Ma trận", "🖨️ 3. Tùy chỉnh & Xuất bản"])

# TAB 1: NHẬP DỮ LIỆU
with tab1:
    st.subheader("Tải lên Ngân hàng câu hỏi của bạn")
    uploaded_file = st.file_uploader("Định dạng .xlsx (Các cột: Chuong, Muc_do, Noi_dung, A, B, C, D, Dap_an_dung)", type=["xlsx"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ Đọc thành công {len(df)} câu hỏi.")
        
        # Thống kê nhanh
        col1, col2, col3, col4 = st.columns(4)
        stats = df['Muc_do'].value_counts()
        col1.metric("Nhận biết (NB)", stats.get('NB', 0))
        col2.metric("Thông hiểu (TH)", stats.get('TH', 0))
        col3.metric("Vận dụng (VD)", stats.get('VD', 0))
        col4.metric("Vận dụng cao (VDC)", stats.get('VDC', 0))
        
        with st.expander("👁️ Xem trước dữ liệu thô"):
            st.dataframe(df.head(10))

# TAB 2: TRỢ LÝ AI
with tab2:
    if not uploaded_file:
        st.warning("⚠️ Vui lòng tải file dữ liệu ở Bước 1 trước.")
    else:
        st.subheader("🤖 Trợ lý AI phân bổ cấu trúc đề")
        st.markdown("Hệ thống sẽ dựa vào dữ liệu khả dụng để tự động chia đều câu hỏi vào các chương.")
        
        col_ai1, col_ai2 = st.columns(2)
        with col_ai1:
            total_questions = st.number_input("Tổng số câu trắc nghiệm muốn tạo:", min_value=1, max_value=200, value=40)
        with col_ai2:
            strategy = st.selectbox("Chiến lược Ma trận", [
                "Chuẩn BGDĐT (40% NB - 30% TH - 20% VD - 10% VDC)",
                "Cơ bản/GDKT (50% NB - 30% TH - 10% VD - 10% VDC)",
                "Phân hóa cao (30% NB - 30% TH - 20% VD - 20% VDC)"
            ])
            
        if st.button("✨ TỰ ĐỘNG PHÂN BỔ BẰNG AI", type="primary"):
            if "Chuẩn" in strategy: ratio = (40, 30, 20, 10)
            elif "Cơ bản" in strategy: ratio = (50, 30, 10, 10)
            else: ratio = (30, 30, 20, 20)
            
            # Chạy thuật toán
            st.session_state.matrix_df = ai_generate_matrix(df, total_questions, ratio)
            st.success("🎉 AI đã tính toán xong! Chuyển sang Bước 3 để xem bảng Đặc tả và Xuất đề.")

# TAB 3: TÙY CHỈNH & XUẤT BẢN
with tab3:
    if not uploaded_file:
        st.warning("⚠️ Vui lòng tải file dữ liệu ở Bước 1 trước.")
    elif st.session_state.matrix_df.empty:
        st.info("👈 Hãy dùng Trợ lý AI ở Bước 2 để tự động tạo khung ma trận trước.")
    else:
        st.subheader("Bảng Ma Trận Đặc Tả (Có thể chỉnh sửa trực tiếp)")
        st.caption("Bạn có thể thay đổi số lượng câu ở từng ô. Cột 'Tổng' sẽ tự động tính.")
        
        # Bảng dữ liệu tương tác
        edited_matrix = st.data_editor(
            st.session_state.matrix_df, 
            use_container_width=True,
            column_config={
                "NB": st.column_config.NumberColumn("Nhận biết", min_value=0),
                "TH": st.column_config.NumberColumn("Thông hiểu", min_value=0),
                "VD": st.column_config.NumberColumn("Vận dụng", min_value=0),
                "VDC": st.column_config.NumberColumn("Vận dụng cao", min_value=0),
            }
        )
        
        # Kiểm tra tính hợp lệ
        is_valid = True
        total_selected = int(edited_matrix.sum().sum())
        st.markdown(f"**Tổng số câu chốt:** `{total_selected}` câu.")
        
        for chapter, row in edited_matrix.iterrows():
            for lvl in ['NB', 'TH', 'VD', 'VDC']:
                req = row[lvl]
                avail = len(df[(df['Chuong'] == chapter) & (df['Muc_do'] == lvl)])
                if req > avail:
                    st.error(f"❌ {chapter} - Mức {lvl}: Yêu cầu {req} câu nhưng ngân hàng chỉ có {avail} câu!")
                    is_valid = False

        st.divider()
        if is_valid:
            if st.button("🚀 BỐC ĐỀ & XUẤT FILE WORD", type="primary", use_container_width=True):
                with st.spinner('Hệ thống đang trộn đề và tạo file Word...'):
                    exam_list = []
                    for chapter, row in edited_matrix.iterrows():
                        for lvl in ['NB', 'TH', 'VD', 'VDC']:
                            count = int(row[lvl])
                            if count > 0:
                                pool = df[(df['Chuong'] == chapter) & (df['Muc_do'] == lvl)]
                                selected = pool.sample(n=count)
                                for _, q_row in selected.iterrows():
                                    exam_list.append(shuffle_question(q_row))
                                    
                    random.shuffle(exam_list)
                    info = {'school': school, 'teacher': teacher, 'subject': subject, 'time': duration}
                    docx_file = export_full_exam(exam_list, info, edited_matrix)
                    
                    st.balloons()
                    st.success("✅ Đã tạo thành công! Vui lòng tải file bên dưới.")
                    st.download_button(
                        label="📥 TẢI ĐỀ KIỂM TRA & MA TRẬN (.docx)",
                        data=docx_file,
                        file_name=f"De_Kiem_Tra_{subject}_Pro.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
        else:
            st.error("Vui lòng chỉnh sửa lại bảng Ma trận hoặc bổ sung câu hỏi vào Excel để có thể Xuất đề.")
