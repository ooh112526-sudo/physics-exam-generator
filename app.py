import streamlit as st
import docx
from docx.shared import Pt
import random
import io
import re

# 設定頁面資訊
st.set_page_config(page_title="物理題庫系統 (Physics Exam Generator)", layout="wide", page_icon="🧲")

# ==========================================
# 核心邏輯類別與函式
# ==========================================

class Question:
    def __init__(self, q_type, content, options=None, answer=None, original_id=0):
        self.id = original_id
        self.type = q_type  # 'Single', 'Multi', 'Fill'
        self.content = content
        self.options = options if options else []  # list of strings
        self.answer = answer  # 'A', 'ABC', or text for fill-in

def parse_docx(file_bytes):
    """解析上傳的 Word 檔案"""
    doc = docx.Document(io.BytesIO(file_bytes))
    questions = []
    current_q = None
    state = None
    # 用於移除選項開頭的 (A) (B) 等標記
    opt_pattern = re.compile(r'^\s*\(?[A-Ea-e]\)?\s*[.、]?\s*')
    q_id_counter = 1

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue

        # 偵測題型標記 [Type:Single]
        if text.startswith('[Type:'):
            if current_q: questions.append(current_q)
            q_type_str = text.split(':')[1].replace(']', '').strip()
            current_q = Question(q_type=q_type_str, content="", options=[], answer="", original_id=q_id_counter)
            q_id_counter += 1
            state = None
            continue

        # 狀態切換
        if text.startswith('[Q]'):
            state = 'Q'; continue
        elif text.startswith('[Opt]'):
            state = 'Opt'; continue
        elif text.startswith('[Ans]'):
            remain_text = text.replace('[Ans]', '').strip()
            if remain_text and current_q: current_q.answer = remain_text
            state = 'Ans'; continue

        # 填入內容
        if current_q:
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
    # 建立索引對照表 (0='A', 1='B'...)
    char_to_idx = {chr(65+i): i for i in range(len(original_opts))}
    
    # 找出正確答案的「內容」
    correct_indices = []
    for char in original_ans:
        if char in char_to_idx: correct_indices.append(char_to_idx[char])
            
    correct_contents = [original_opts[i] for i in correct_indices]
    
    # 打亂選項
    shuffled_opts_data = list(enumerate(original_opts))
    random.shuffle(shuffled_opts_data)
    new_options = [data[1] for data in shuffled_opts_data]
    
    # 尋找正確答案的新位置
    new_ans_chars = []
    for content in correct_contents:
        try:
            new_idx = new_options.index(content)
            new_ans_chars.append(chr(65 + new_idx))
        except ValueError: pass
            
    new_ans_chars.sort()
    new_answer_str = "".join(new_ans_chars)

    # 回傳一個新的物件，確保不修改原始題庫
    new_q = Question(question.type, question.content, new_options, new_answer_str, question.id)
    return new_q

def generate_word_files(selected_questions, shuffle=True):
    """生成 Word 試卷與詳解"""
    exam_doc = docx.Document()
    ans_doc = docx.Document()
    
    # 設定基本字型
    style = exam_doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # 標題
    exam_doc.add_heading('物理科 試題卷', 0)
    ans_doc.add_heading('物理科 答案卷', 0)
    exam_doc.add_paragraph('班級：__________  姓名：__________  座號：__________\n')
    
    for idx, q in enumerate(selected_questions, 1):
        processed_q = q
        # 若啟用亂數且非填充題，則進行重排
        if shuffle and q.type in ['Single', 'Multi']:
            processed_q = shuffle_options_and_update_answer(q)
        
        # --- 寫入試題卷 ---
        p = exam_doc.add_paragraph()
        q_type_text = {'Single': '單選', 'Multi': '多選', 'Fill': '填充'}.get(q.type, '未知')
        runner = p.add_run(f"{idx}. ({q_type_text}) {processed_q.content.strip()}")
        runner.bold = True
        
        if q.type != 'Fill':
            for i, opt in enumerate(processed_q.options):
                exam_doc.add_paragraph(f"({chr(65+i)}) {opt}")
        else:
            exam_doc.add_paragraph("______________________")
        exam_doc.add_paragraph("") # 空行分隔
        
        # --- 寫入答案卷 ---
        ans_p = ans_doc.add_paragraph()
        ans_p.add_run(f"{idx}. ").bold = True
        ans_p.add_run(f"{processed_q.answer}")

    # 儲存到記憶體
    exam_io = io.BytesIO()
    ans_io = io.BytesIO()
    exam_doc.save(exam_io)
    ans_doc.save(ans_io)
    exam_io.seek(0)
    ans_io.seek(0)
    return exam_io, ans_io

# ==========================================
# Session State 初始化 (用於暫存題目)
# ==========================================
if 'question_pool' not in st.session_state:
    st.session_state['question_pool'] = []

# ==========================================
# Streamlit 介面
# ==========================================

st.title("🧲 物理題庫自動組卷系統 v2.0")
st.markdown("支援 **手動輸入** 與 **Word 匯入** 混合出題模式。")

# --- 側邊欄：管理題庫 ---
with st.sidebar:
    st.header("📦 題庫管理")
    count = len(st.session_state['question_pool'])
    st.metric("目前題庫總數", f"{count} 題")
    
    if count > 0:
        if st.button("🗑️ 清空所有題目", type="primary"):
            st.session_state['question_pool'] = []
            st.rerun()
    
    st.divider()
    st.info("💡 提示：您可以先從 Word 匯入題庫，再手動補充幾題，最後一起匯出。")
    
    # 範本下載
    sample_doc = docx.Document()
    sample_doc.add_paragraph("[Type:Single]\n[Q]\n(範例) 雙狹縫干涉實驗中...\n[Opt]\n(A)變大\n(B)變小\n[Ans] A")
    sample_io = io.BytesIO()
    sample_doc.save(sample_io)
    sample_io.seek(0)
    st.download_button("📥 下載 Word 匯入格式範本", sample_io, "template.docx")

# --- 主畫面分頁 ---
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
            new_q = Question(new_q_type, new_q_content, new_q_options, new_q_ans, q_id)
            st.session_state['question_pool'].append(new_q)
            st.success("題目已加入！請切換到「選題與匯出」查看。")

# === Tab 2: Word 匯入 ===
with tab2:
    st.subheader("批次匯入題目")
    st.write("請依照側邊欄的範本格式準備 Word 檔。")
    uploaded_file = st.file_uploader("上傳 Word (.docx) 檔案", type=['docx'])
    
    if uploaded_file:
        if st.button("解析並加入題庫"):
            try:
                imported_qs = parse_docx(uploaded_file.read())
                if imported_qs:
                    st.session_state['question_pool'].extend(imported_qs)
                    st.success(f"成功匯入 {len(imported_qs)} 題！目前總數：{len(st.session_state['question_pool'])}")
                else:
                    st.warning("檔案中未偵測到符合格式的題目，請檢查標籤是否正確 (如 [Type:Single], [Q]...)")
            except Exception as e:
                st.error(f"解析失敗：{e}")

# === Tab 3: 選題與匯出 ===
with tab3:
    st.subheader("預覽與組卷")
    
    if not st.session_state['question_pool']:
        st.info("目前題庫是空的，請先到前兩個分頁新增題目。")
    else:
        # 全選控制
        col_ctrl, _ = st.columns([2, 8])
        with col_ctrl:
            select_all = st.checkbox("全選所有題目", value=True)
        
        selected_indices = []
        st.write("---")
        
        # 顯示題目列表
        for i, q in enumerate(st.session_state['question_pool']):
            col_check, col_text = st.columns([0.5, 9.5])
            with col_check:
                # 勾選框
                is_checked = st.checkbox("選取", value=select_all, key=f"sel_{i}", label_visibility="collapsed")
                if is_checked:
                    selected_indices.append(i)
            
            with col_text:
                type_badge = {'Single': '🟢單選', 'Multi': '🔵多選', 'Fill': '🟠填充'}.get(q.type)
                # 使用 Expander 收折題目詳情
                with st.expander(f"{i+1}. {type_badge} {q.content.splitlines()[0][:40]}..."):
                    st.markdown(f"**題目**：\n{q.content}")
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
        
        do_shuffle = st.checkbox("啟用選項亂數重排 (Shuffle Options)", value=True, help="自動打亂選項順序並修正答案")
        
        if st.button("🚀 生成 Word 試卷", type="primary", disabled=len(selected_indices)==0):
            # 取得選中的題目
            final_qs = [st.session_state['question_pool'][i] for i in selected_indices]
            # 生成檔案
            exam_file, ans_file = generate_word_files(final_qs, shuffle=do_shuffle)
            
            st.success("檔案生成完畢！請點擊下方按鈕下載：")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button("📄 下載試題卷", exam_file, "物理試題卷.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_d2:
                st.download_button("🔑 下載詳解卷", ans_file, "物理詳解卷.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")