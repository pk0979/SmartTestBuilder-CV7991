import streamlit as st
import pandas as pd
import random
import io
import json
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="EduTest Pro - CV7991", page_icon="🎓", layout="wide")

st.markdown("""
<style>
    .stMetric { background-color: #f0f2f6; padding: 10px; border-radius: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 4px 4px 0 0; padding-top: 10px; padding-bottom: 10px; }
    .stTabs [aria-selected="true"] { background-color: #e6f0ff; border-bottom: 3px solid #0066cc; font-weight: bold;}
</style>
""", unsafe_allow_html=True)

# --- KHỞI TẠO SESSION STATE ---
if 'db_df' not in st.session_state:
    st.session_state.db_df = pd.DataFrame(columns=['Chuong', 'Muc_do', 'Noi_dung', 'A', 'B', 'C', 'D', 'Dap_an_dung'])
if 'matrix_df' not in st.session_state:
    st.session_state.matrix_df = pd.DataFrame()

# --- HÀM TIỆN ÍCH ---
def style_text(paragraph, text, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    return run

# --- HÀM GỌI GEMINI AI TẠO CÂU HỎI ---
def generate_questions_with_ai(api_key, subject, chapter, nb, th, vd, vdc):
    prompt = f"""
    Bạn là một chuyên gia giáo dục tại Việt Nam. Hãy soạn câu hỏi trắc nghiệm môn {subject}, phần/chủ đề "{chapter}" bám sát yêu cầu Công văn 7991/BGDĐT-GDTrH.
    Số lượng: {nb} câu Nhận biết (NB), {th} câu Thông hiểu (TH), {vd} câu Vận dụng (VD), {vdc} câu Vận dụng cao (VDC).
    Yêu cầu:
    - Câu hỏi rõ ràng, khoa học.
    - 4 phương án A, B, C, D hợp lý, không quá chênh lệch độ dài.
    Trả về kết quả CHUẨN JSON (Array of Objects) với các key sau. (TUYỆT ĐỐI KHÔNG chứa ký tự markdown như ```json):
    [
        {{"Chuong": "{chapter}", "Muc_do": "NB", "Noi_dung": "...", "A": "...", "B": "...", "C": "...", "D": "...", "Dap_an_dung": "A"}},
        ...
    ]
    """
    
    headers = {'Content-Type': 'application/json'}
    data = {
        "contents": [{"parts":[{"text": prompt}]}],
        "generationConfig": {"response_mime_type": "application/json"}
    }
    
    # Danh sách các phiên bản AI dự phòng (từ mới nhất đến cũ hơn)
    # Hệ thống sẽ tự quét tìm phiên bản nào phù hợp với API Key của bạn
    models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-flash-8b", "gemini-1.5-pro"]
    
    for model in models:
        url = f"[https://generativelanguage.googleapis.com/v1beta/models/](https://generativelanguage.googleapis.com/v1beta/models/){model}:generateContent?key={api_key}"
        try:
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 200:
                result = response.json()['candidates'][0]['content']['parts'][0]['text']
                result = result.replace('```json', '').replace('```', '').strip()
                return json.loads(result)
            elif response.status_code == 404:
                continue # Nếu bị báo lỗi 404 (Không tìm thấy model), tự động bỏ qua và thử model tiếp theo
            else:
                st.error(f"Lỗi API ({model}): {response.text}")
                return None
        except Exception as e:
            continue # Bỏ qua lỗi kết nối mạng tạm thời
            
    st.error("❌ Lỗi: API Key của bạn không hỗ trợ các model Gemini hiện tại. Hãy tạo Project mới trên Google AI Studio.")
    return None

# --- HÀM AI PHÂN BỔ MA TRẬN ---
def ai_generate_matrix(df, total_q, strategy_ratio):
    chapters = df['Chuong'].unique()
    levels = ['NB', 'TH', 'VD', 'VDC']
    targets = {lvl: int(total_q * r / 100) for lvl, r in zip(levels, strategy_ratio)}
    targets['VDC'] = total_q - sum([targets[l] for l in ['NB', 'TH', 'VD']])
    
    plan = {ch: {lvl: 0 for lvl in levels} for ch in chapters}
    available = {ch: {lvl: len(df[(df['Chuong'] == ch) & (df['Muc_do'] == lvl)]) for lvl in levels} for ch in chapters}
    
    for lvl in levels:
        t = targets[lvl]
        while t > 0:
            total_avail_lvl = sum([available[ch][lvl] for ch in chapters])
            if total_avail_lvl == 0:
                st.warning(f"⚠️ Ngân hàng thiếu câu hỏi mức {lvl}! Đã phân bổ được {targets[lvl] - t}/{targets[lvl]} câu.")
                break
                
            for ch in chapters:
                if t > 0 and available[ch][lvl] > 0:
                    plan[ch][lvl] += 1
                    available[ch][lvl] -= 1
                    t -= 1
                    
    return pd.DataFrame.from_dict(plan, orient='index')

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
st.title("🎓 Hệ Thống Tạo Đề & Sinh Câu Hỏi AI (Chuẩn CV 7991)")

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3074/3074058.png", width=100)
    st.header("⚙️ Thông tin chung")
    school = st.text_input("Tên trường", "THCS Đa Phước")
    teacher = st.text_input("Tên giáo viên", "Tống Phước Khải")
    subject = st.text_input("Môn học", "Toán")
    duration = st.number_input("Thời gian (phút)", value=90)
    
    st.divider()
    st.markdown("🔑 **Khóa API AI (Tùy chọn)**")
    api_key = st.text_input("Nhập Gemini API Key", type="password", help="Dùng để sinh câu hỏi tự động.")
    st.caption("Lấy key miễn phí tại: [Google AI Studio](https://aistudio.google.com/)")

# Quy trình 4 bước
tab1, tab2, tab3, tab4 = st.tabs(["📁 1. Dữ liệu", "✨ 2. Sinh Câu Hỏi AI", "🤖 3. Lập Ma Trận", "🖨️ 4. Xuất Bản"])

# TAB 1: NHẬP DỮ LIỆU TỪ EXCEL
with tab1:
    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("Tải lên Ngân hàng câu hỏi")
        uploaded_file = st.file_uploader("Định dạng .xlsx (Cột: Chuong, Muc_do, Noi_dung, A, B, C, D, Dap_an_dung)", type=["xlsx"])
        if st.button("Tải vào hệ thống", type="primary"):
            if uploaded_file:
                new_df = pd.read_excel(uploaded_file)
                st.session_state.db_df = pd.concat([st.session_state.db_df, new_df], ignore_index=True).drop_duplicates()
                st.success("Đã bổ sung câu hỏi vào Kho dữ liệu chung!")
    with col2:
        st.subheader("Thống kê Kho Dữ Liệu Hiện Tại")
        st.info(f"📚 Tổng số câu hỏi trong bộ nhớ: **{len(st.session_state.db_df)}** câu")
        if not st.session_state.db_df.empty:
            stats = st.session_state.db_df['Muc_do'].value_counts()
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("NB", stats.get('NB', 0))
            c2.metric("TH", stats.get('TH', 0))
            c3.metric("VD", stats.get('VD', 0))
            c4.metric("VDC", stats.get('VDC', 0))
            if st.button("🗑️ Xóa sạch bộ nhớ", type="secondary"):
                st.session_state.db_df = pd.DataFrame(columns=['Chuong', 'Muc_do', 'Noi_dung', 'A', 'B', 'C', 'D', 'Dap_an_dung'])
                st.rerun()

# TAB 2: AI TẠO CÂU HỎI MỚI
with tab2:
    st.subheader("Trợ lý AI - Tạo mới câu hỏi lấp đầy ma trận")
    st.markdown("Nếu ngân hàng của bạn thiếu câu hỏi (Đặc biệt là mức độ Vận dụng cao), hãy yêu cầu AI tạo ngay lập tức.")
    
    col2_1, col2_2 = st.columns(2)
    with col2_1:
        chapter_ai = st.text_input("Nhập tên Chủ đề / Chương (VD: Đạo hàm, Rễ cây...):")
        
    with col2_2:
        col_nb, col_th, col_vd, col_vdc = st.columns(4)
        ai_nb = col_nb.number_input("Số câu NB", min_value=0, value=2)
        ai_th = col_th.number_input("Số câu TH", min_value=0, value=2)
        ai_vd = col_vd.number_input("Số câu VD", min_value=0, value=1)
        ai_vdc = col_vdc.number_input("Số câu VDC", min_value=0, value=1)
        
    if st.button("✨ XUẤT XƯỞNG CÂU HỎI MỚI", type="primary"):
        if not api_key:
            st.error("❌ Bạn cần nhập 'Gemini API Key' ở thanh Menu bên trái trước.")
        elif not chapter_ai:
            st.error("❌ Vui lòng nhập tên Chủ đề/Chương cần tạo.")
        else:
            with st.spinner('🤖 AI đang suy nghĩ và biên soạn câu hỏi... (có thể mất 10-20 giây)'):
                gen_data = generate_questions_with_ai(api_key, subject, chapter_ai, ai_nb, ai_th, ai_vd, ai_vdc)
                if gen_data:
                    new_ai_df = pd.DataFrame(gen_data)
                    st.session_state.db_df = pd.concat([st.session_state.db_df, new_ai_df], ignore_index=True)
                    st.success("🎉 Đã tạo thành công và nạp vào Kho Dữ Liệu! Bạn có thể xem kết quả bên dưới.")
                    st.dataframe(new_ai_df)

# TAB 3: AI PHÂN BỔ MA TRẬN
with tab3:
    if st.session_state.db_df.empty:
        st.warning("⚠️ Kho dữ liệu đang trống. Hãy qua Tab 1 hoặc Tab 2 để thêm câu hỏi trước.")
    else:
        st.subheader("🤖 AI Tính toán Cấu trúc đề (Ma trận)")
        col_ai1, col_ai2 = st.columns(2)
        with col_ai1:
            total_questions = st.number_input("Tổng số câu trắc nghiệm đề thi:", min_value=1, max_value=200, value=40)
        with col_ai2:
            strategy = st.selectbox("Chiến lược Ma trận CV7991", [
                "Chuẩn BGDĐT (40% NB - 30% TH - 20% VD - 10% VDC)",
                "Cơ bản (50% NB - 30% TH - 10% VD - 10% VDC)",
                "Nâng cao (30% NB - 30% TH - 20% VD - 20% VDC)"
            ])
            
        if st.button("✨ TỰ ĐỘNG LẬP KHUNG MA TRẬN", type="primary"):
            if "Chuẩn" in strategy: ratio = (40, 30, 20, 10)
            elif "Cơ bản" in strategy: ratio = (50, 30, 10, 10)
            else: ratio = (30, 30, 20, 20)
            
            st.session_state.matrix_df = ai_generate_matrix(st.session_state.db_df, total_questions, ratio)
            st.success("🎉 Đã phân bổ xong! Chuyển sang Tab 4 để xuất đề.")

# TAB 4: TÙY CHỈNH & XUẤT BẢN
with tab4:
    if st.session_state.matrix_df.empty:
        st.info("👈 Hãy dùng Tab 3 để tự động tạo khung ma trận trước.")
    else:
        st.subheader("Bảng Ma Trận Đặc Tả (Có thể tinh chỉnh)")
        
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
        
        is_valid = True
        total_selected = int(edited_matrix.sum().sum())
        st.markdown(f"**Tổng số câu chốt:** `{total_selected}` câu.")
        
        for chapter, row in edited_matrix.iterrows():
            for lvl in ['NB', 'TH', 'VD', 'VDC']:
                req = row[lvl]
                avail = len(st.session_state.db_df[(st.session_state.db_df['Chuong'] == chapter) & (st.session_state.db_df['Muc_do'] == lvl)])
                if req > avail:
                    st.error(f"❌ Chương '{chapter}' - Mức {lvl}: Cần {req} câu nhưng kho chỉ có {avail} câu! (Mẹo: Qua Tab 2 nhờ AI tạo thêm)")
                    is_valid = False

        st.divider()
        if is_valid:
            if st.button("🚀 BỐC ĐỀ & XUẤT FILE WORD", type="primary", use_container_width=True):
                with st.spinner('Đang trộn phương án và tạo file...'):
                    exam_list = []
                    for chapter, row in edited_matrix.iterrows():
                        for lvl in ['NB', 'TH', 'VD', 'VDC']:
                            count = int(row[lvl])
                            if count > 0:
                                pool = st.session_state.db_df[(st.session_state.db_df['Chuong'] == chapter) & (st.session_state.db_df['Muc_do'] == lvl)]
                                selected = pool.sample(n=count)
                                for _, q_row in selected.iterrows():
                                    exam_list.append(shuffle_question(q_row))
                                    
                    random.shuffle(exam_list)
                    info = {'school': school, 'teacher': teacher, 'subject': subject, 'time': duration}
                    docx_file = export_full_exam(exam_list, info, edited_matrix)
                    
                    st.balloons()
                    st.success("✅ Đã tạo thành công! Hãy tải file bên dưới.")
                    st.download_button(
                        label="📥 TẢI ĐỀ KIỂM TRA & MA TRẬN (.docx)",
                        data=docx_file,
                        file_name=f"De_Kiem_Tra_{subject}_AI_CV7991.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
