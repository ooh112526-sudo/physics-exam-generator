import streamlit as st
import docx
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
import random
import io
import re

# 設定頁面資訊
st.set_page_config(page_title="物理題庫系統 (Physics Exam Generator)", layout="wide", page_icon="🧲")

# ==========================================
# 核心邏輯類別與函式
# ==========================================

class Question:
    def __init__(self, q_type, content, options=None, answer=None, original_id=0, image_data=None):
        self.id = original_id
        self.type = q_type  # 'Single', 'Multi', 'Fill'
        self.content = content
        self.options = options if options else []  # list of strings
        self.answer = answer  # 'A', 'ABC', or text for fill-in
        self.image_data = image_data  # BytesIO or bytes object

def extract_images_from_paragraph(paragraph, doc_part):
    """
    從 Word 段落中擷取圖片 (Blob data)
    這是比較進階的寫法，直接從 XML 尋找關聯的圖片 ID
    """
    images = []
    # XML Namespace map
    nsmap = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    # 尋找所有 blip 元素 (圖片參照點)
    # paragraph._element 是 lxml 的 element
    blips = paragraph._element.findall('.//a:blip', namespaces=nsmap)
    
    for blip in blips:
        # 取得 rId (Relationship ID)
        embed_attr = blip.get(f"{{{nsmap['r']}}}embed")
        if embed_attr and embed_attr in doc_part.rels:
            part = doc_part.rels[embed_attr].target_part
            # 確認是圖片 Part
            if "image" in part.content_type:
                images.append(part.blob)
    return images

def parse_docx(file_bytes):
    """解析上傳的 Word 檔案 (含圖片擷取)"""
    doc = docx.Document(io.BytesIO(file_bytes))
    doc_part = doc.part # 取得 Document Part 以存取 Relationships
    
    questions = []
    current_q = None
    state = None
    opt_pattern = re.compile(r'^\s*\(?[A-Ea-e]\)?\s*[.、]?\s*')
    q_id_counter = 1

    for para in doc.paragraphs:
        text = para.text.strip()
        
        # 嘗試從該段落擷取圖片
        found_images = extract_images_from_paragraph(para, doc_part)
        
        # 1. 偵測新題目
        if text.startswith('[Type:'):
            if current_q: questions.append(current_q)
            q_type_str = text.split(':')[1].replace(']', '').strip()
            current_q = Question(q_type=q_type_str, content="", options=[], answer="", original_id=q_id_counter)
            q_id_counter += 1
            state = None
            continue

        # 2. 狀態切換
        if text.startswith('[Q]'):
            state = 'Q'; continue
        elif text.startswith('[Opt]'):
            state = 'Opt'; continue
        elif text.startswith('[Ans]'):
            remain_text = text.replace('[Ans]', '').strip()
            if remain_text and current_q: current_q.answer = remain_text
            state = 'Ans'; continue

        # 3. 填入內容與圖片
        if current_q:
            # 如果這段落有圖片，且目前是題目區塊，則加入圖片
            # (目前簡化邏輯：一題只存一張主要圖片，若有多張會覆蓋，可視需求調整)
            if found_images and state == 'Q':
                current_q.image_data = found_images[0]

            if not text: continue # 忽略純空行(但前面有檢查圖片，所以圖片行如果是空的文字也會被處理)

            if state == 'Q': current_q.content += text + "\n"
            elif state == 'Opt':
                clean_opt = opt_pattern.sub('', text)
                current_q.options.append(clean_opt)
            elif state == 'Ans': current_q.answer += text

    if current_q: questions.append(current_q)
    return questions

def shuffle_options_and_update_answer(question):
    """核心演算法：打亂選項並修正答案索引"""
    if question.type == 'Fill': return question

    original_opts = question.options
    original_ans = question.answer.strip().upper()
    char_to_idx = {chr(65+i): i for i in range(len(original_opts))}
    
    correct_indices = []
    for char in original_ans:
        if char in char_to_idx: correct_indices.append(char_to_idx[char])
            
    correct_contents = [original_opts[i] for i in correct_indices]
    
    shuffled_opts_data = list(enumerate(original_opts))
    random.shuffle(shuffled_opts_data)
    new_options = [data[1] for data in shuffled_opts_data]
    
    new_ans_chars = []
    for content in correct_contents:
        try:
            new_idx = new_options.index(content)
            new_ans_chars.append(chr(65 + new_idx))
        except ValueError: pass
            
    new_ans_chars.sort()
    new_answer_str = "".join(new_ans_chars)

    # 包含 image_data 一起複製
    new_q = Question(question.type, question.content, new_options, new_answer_str, question.id, question.image_data)
    return new_q

def generate_word_files(selected_questions, shuffle=True):
    """生成 Word 試卷與詳解 (含圖片)"""
    exam_doc = docx.Document()
    ans_doc = docx.Document()
    
    style = exam_doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    exam_doc.add_heading('物理科 試題卷', 0)
    ans_doc.add_heading('物理科 答案卷', 0)
    exam_doc.add_paragraph('班級：__________  姓名：__________  座號：__________\n')
    
    for idx, q in enumerate(selected_questions, 1):
        processed_q = q
        if shuffle and q.type in ['Single', 'Multi']:
            processed_q = shuffle_options_and_update_answer(q)
        
        # --- 試題卷 ---
        p = exam_doc.add_paragraph()
        q_type_text = {'Single': '單選', 'Multi': '多選', 'Fill': '填充'}.get(q.type, '未知')
        runner = p.add_run(f"{idx}. ({q_type_text}) {processed_q.content.strip()}")
        runner.bold = True
        
        # 插入圖片 (如果有)
        if processed_q.image_data:
            try:
                # 需將 bytes 轉為 stream
                img_stream = io.BytesIO(processed_q.image_data)
                # 預設寬度 3 英吋，可自行調整
                exam_doc.add_picture(img_stream, width=Inches(3.0))
            except Exception as e:
                print(f"Error adding picture: {e}")

        if q.type != 'Fill':
            for i, opt in enumerate(processed_q.options):
                exam_doc.add_paragraph(f"({chr(65+i)}) {opt}")
        else:
            exam_doc.add_paragraph("______________________")
        exam_doc.add_paragraph("") 
        
        # --- 答案卷 ---
        ans_p = ans_doc.add_paragraph()
        ans_p.add_run(f"{idx}. ").bold = True
        ans_p.add_run(f"{processed_q.answer}")

    exam_io = io.BytesIO()
    ans_io = io.BytesIO()
    exam_doc.save(exam_io)
    ans_doc.save(ans_io)
    exam_io.seek(0)
    ans_io.seek(0)
    return exam_io, ans_io

# ==========================================
# Session State
# ==========================================
if 'question_pool' not in st.session_state:
    st.session_state['question_pool'] = []

# ==========================================
# Streamlit 介面
# ==========================================

st.title("🧲 物理題庫自動組卷系統 v2.5 (含圖片支援)")
st.markdown("支援 **手動輸入(含圖片)** 與 **Word 匯入(自動抓圖)** 混合出題模式。")

# --- 側邊欄 ---
with st.sidebar:
    st.header("📦 題庫管理")
    count = len(st.session_state['question_pool'])
    st.metric("目前題庫總數", f"{count} 題")
    
    if count > 0:
        if st.button("🗑️ 清空所有題目", type="primary"):
            st.session_state['question_pool'] = []
            st.rerun()
    
    st.divider()
    st.info("💡 提示：Word 匯入時，程式會嘗試抓取 `[Q]` 區塊內的圖片。手動輸入時可直接上傳圖片檔。")
    
    # 範本下載 (簡單文字版，圖片建議手動測試)
    sample_doc = docx.Document()
    sample_doc.add_paragraph("[Type:Single]\n[Q]\n(範例) 下圖為波動示意圖...\n(請在此插入圖片)\n[Opt]\n(A)變大\n(B)變小\n[Ans] A")
    sample_io = io.BytesIO()
    sample_doc.save(sample_io)
    sample_io.seek(0)
    st.download_button("📥 下載 Word 範本", sample_io, "template.docx")

# --- 主畫面 ---
tab1, tab2, tab3 = st.tabs(["✍️ 手動新增題目", "📁 從 Word 匯入", "🚀 選題與匯出"])

# === Tab 1: 手動輸入 ===
with tab1:
    st.subheader("新增單一題目")
    
    c1, c2 = st.columns([1, 3])
    with c1:
        new_q_type = st.selectbox("題型", ["Single", "Multi", "Fill"], format_func=lambda x: {'Single':'單選題', 'Multi':'多選題', 'Fill':'填充題'}[x])
    with c2:
        new_q_ans = st.text_input("正確答案", placeholder="選擇題填代號(如 A, AC)，填充題填文字")

    new_q_content = st.text_area("題目內容", height=100, placeholder="請輸入題目敘述...")
    
    # 圖片上傳區
    new_q_image = st.file_uploader("上傳圖片 (選用)", type=['png', 'jpg', 'jpeg'], help="若題目包含電路圖或示意圖請在此上傳")
    
    new_q_options = []
    if new_q_type in ["Single", "Multi"]:
        opts_text = st.text_area("選項 (每一行一個選項)", height=150, placeholder="1.5 倍\n0.67 倍\n2.25 倍\n不變")
        if opts_text:
            new_q_options = [line.strip() for line in opts_text.split('\n') if line.strip()]

    if st.button("➕ 加入題庫", type="secondary"):
        if not new_q_content:
            st.error("請輸入題目內容")
        elif new_q_type != 'Fill' and not new_q_options:
            st.error("選擇題必須提供選項")
        else:
            q_id = len(st.session_state['question_pool']) + 1
            
            # 處理圖片
            img_bytes = None
            if new_q_image is not None:
                img_bytes = new_q_image.getvalue()

            new_q = Question(new_q_type, new_q_content, new_q_options, new_q_ans, q_id, image_data=img_bytes)
            st.session_state['question_pool'].append(new_q)
            st.success("題目(含圖片)已加入！")

# === Tab 2: Word 匯入 ===
with tab2:
    st.subheader("批次匯入題目")
    st.write("請依照範本格式準備 Word 檔。若題目段落中有插入圖片，系統會嘗試自動擷取。")
    uploaded_file = st.file_uploader("上傳 Word (.docx) 檔案", type=['docx'])
    
    if uploaded_file:
        if st.button("解析並加入題庫"):
            try:
                imported_qs = parse_docx(uploaded_file.read())
                if imported_qs:
                    st.session_state['question_pool'].extend(imported_qs)
                    st.success(f"成功匯入 {len(imported_qs)} 題！")
                else:
                    st.warning("未偵測到題目，請檢查格式標籤。")
            except Exception as e:
                st.error(f"解析失敗：{e}")

# === Tab 3: 選題與匯出 ===
with tab3:
    st.subheader("預覽與組卷")
    
    if not st.session_state['question_pool']:
        st.info("目前題庫是空的。")
    else:
        col_ctrl, _ = st.columns([2, 8])
        with col_ctrl:
            select_all = st.checkbox("全選所有題目", value=True)
        
        selected_indices = []
        st.write("---")
        
        for i, q in enumerate(st.session_state['question_pool']):
            col_check, col_text = st.columns([0.5, 9.5])
            with col_check:
                is_checked = st.checkbox("選取", value=select_all, key=f"sel_{i}", label_visibility="collapsed")
                if is_checked:
                    selected_indices.append(i)
            
            with col_text:
                type_badge = {'Single': '🟢單選', 'Multi': '🔵多選', 'Fill': '🟠填充'}.get(q.type)
                with st.expander(f"{i+1}. {type_badge} {q.content.splitlines()[0][:40]}..."):
                    st.markdown(f"**題目**：\n{q.content}")
                    
                    # 預覽圖片
                    if q.image_data:
                        st.image(q.image_data, caption="題目附圖", width=300)
                        
                    if q.options:
                        st.markdown("**選項**：")
                        for idx, opt in enumerate(q.options):
                            st.text(f"({chr(65+idx)}) {opt}")
                    st.markdown(f"**答案**：`{q.answer}`")
                    
                    if st.button("🗑️ 刪除此題", key=f"del_{i}"):
                        st.session_state['question_pool'].pop(i)
                        st.rerun()

        st.divider()
        st.subheader("匯出設定")
        st.write(f"已選擇: **{len(selected_indices)}** 題")
        
        do_shuffle = st.checkbox("啟用選項亂數重排", value=True)
        
        if st.button("🚀 生成 Word 試卷", type="primary", disabled=len(selected_indices)==0):
            final_qs = [st.session_state['question_pool'][i] for i in selected_indices]
            exam_file, ans_file = generate_word_files(final_qs, shuffle=do_shuffle)
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button("📄 下載試題卷", exam_file, "物理試題卷.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_d2:
                st.download_button("🔑 下載詳解卷", ans_file, "物理詳解卷.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
