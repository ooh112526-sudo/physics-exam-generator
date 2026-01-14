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
# 常數定義：章節與單元資料
# ==========================================

SOURCES = ["一般試題", "學測題", "北模", "全模", "中模"]

PHYSICS_CHAPTERS = {
    "第一章.科學的態度與方法": [
        "1-1 科學的態度", "1-2 科學的方法", "1-3 國際單位制", "1-4 物理學簡介"
    ],
    "第二章.物體的運動": [
        "2-1 物體的運動", "2-2 牛頓三大運動定律", "2-3 生活中常見的力", "2-4 天體運動"
    ],
    "第三章. 物質的組成與交互作用": [
        "3-1 物質的組成", "3-2 原子的結構", "3-3 基本交互作用"
    ],
    "第四章.電與磁的統一": [
        "4-1 電流磁效應", "4-2 電磁感應", "4-3 電與磁的整合", "4-4 光波的特性", "4-5 都卜勒效應"
    ],
    "第五章. 能　量": [
        "5-1 能量的形式", "5-2 微觀尺度下的能量", "5-3 能量守恆", "5-4 質能互換"
    ],
    "第六章.量子現象": [
        "6-1 量子論的誕生", "6-2 光的粒子性", "6-3 物質的波動性", "6-4 波粒二象性", "6-5 原子光譜"
    ]
}

# ==========================================
# 核心邏輯類別與函式
# ==========================================

class Question:
    def __init__(self, q_type, content, options=None, answer=None, original_id=0, image_data=None, 
                 source="一般試題", chapter="", unit=""):
        self.id = original_id
        self.type = q_type  # 'Single', 'Multi', 'Fill'
        self.source = source
        self.chapter = chapter
        self.unit = unit
        self.content = content
        self.options = options if options else []
        self.answer = answer
        self.image_data = image_data

def extract_images_from_paragraph(paragraph, doc_part):
    """從 Word 段落中擷取圖片"""
    images = []
    nsmap = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    blips = paragraph._element.findall('.//a:blip', namespaces=nsmap)
    for blip in blips:
        embed_attr = blip.get(f"{{{nsmap['r']}}}embed")
        if embed_attr and embed_attr in doc_part.rels:
            part = doc_part.rels[embed_attr].target_part
            if "image" in part.content_type:
                images.append(part.blob)
    return images

def parse_docx(file_bytes):
    """解析 Word 檔案 (支援 Source, Chapter, Unit 標籤)"""
    doc = docx.Document(io.BytesIO(file_bytes))
    doc_part = doc.part
    
    questions = []
    current_q = None
    state = None
    opt_pattern = re.compile(r'^\s*\(?[A-Ea-e]\)?\s*[.、]?\s*')
    q_id_counter = 1

    # 預設狀態 (會延續到下一題)
    curr_src = "一般試題"
    curr_chap = ""
    curr_unit = ""

    for para in doc.paragraphs:
        text = para.text.strip()
        found_images = extract_images_from_paragraph(para, doc_part)
        
        # 0. 偵測分類標籤
        if text.startswith('[Src:'):
            curr_src = text.split(':')[1].replace(']', '').strip()
            continue
        if text.startswith('[Chap:'):
            curr_chap = text.split(':')[1].replace(']', '').strip()
            continue
        if text.startswith('[Unit:'):
            curr_unit = text.split(':')[1].replace(']', '').strip()
            continue
        # 相容舊版 [Cat:] 標籤 (視為章節或單元)
        if text.startswith('[Cat:'):
            curr_unit = text.split(':')[1].replace(']', '').strip()
            continue

        # 1. 偵測新題目
        if text.startswith('[Type:'):
            if current_q: questions.append(current_q)
            q_type_str = text.split(':')[1].replace(']', '').strip()
            current_q = Question(
                q_type=q_type_str, 
                content="", 
                options=[], 
                answer="", 
                original_id=q_id_counter, 
                source=curr_src,
                chapter=curr_chap,
                unit=curr_unit
            )
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

        # 3. 填入內容
        if current_q:
            if found_images and state == 'Q':
                current_q.image_data = found_images[0]

            if not text: continue

            if state == 'Q': current_q.content += text + "\n"
            elif state == 'Opt':
                clean_opt = opt_pattern.sub('', text)
                current_q.options.append(clean_opt)
            elif state == 'Ans': current_q.answer += text

    if current_q: questions.append(current_q)
    return questions

def shuffle_options_and_update_answer(question):
    """打亂選項並修正答案"""
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

    return Question(
        question.type, question.content, new_options, new_answer_str, 
        question.id, question.image_data, 
        question.source, question.chapter, question.unit
    )

def generate_word_files(selected_questions, shuffle=True):
    """生成 Word 試卷"""
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
        
        if processed_q.image_data:
            try:
                img_stream = io.BytesIO(processed_q.image_data)
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
        
        # 在詳解卷顯示完整分類資訊
        meta_info = []
        if processed_q.source and processed_q.source != "一般試題": meta_info.append(processed_q.source)
        if processed_q.unit: meta_info.append(processed_q.unit)
        elif processed_q.chapter: meta_info.append(processed_q.chapter)
            
        if meta_info:
            ans_p.add_run(f"  [{' / '.join(meta_info)}]").italic = True

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

st.title("🧲 物理題庫自動組卷系統 v2.7")
st.markdown("支援 **完整章節分類**、**學測/模考來源標記** 與 **圖片功能**。")

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
    st.markdown("""
    **Word 匯入標籤說明：**
    - `[Src:學測題]` 來源
    - `[Chap:第一章...]` 章節
    - `[Unit:1-1...]` 單元
    - `[Type:Single]` 題型
    """)
    
    sample_doc = docx.Document()
    sample_doc.add_paragraph("[Src:北模]")
    sample_doc.add_paragraph("[Chap:第四章.電與磁的統一]")
    sample_doc.add_paragraph("[Unit:4-1 電流磁效應]")
    sample_doc.add_paragraph("[Type:Single]\n[Q]\n(範例) 下列關於安培右手定則...\n[Opt]\n(A)選項一\n(B)選項二\n[Ans] A")
    sample_io = io.BytesIO()
    sample_doc.save(sample_io)
    sample_io.seek(0)
    st.download_button("📥 下載 Word 範本", sample_io, "template.docx")

# --- 主畫面 ---
tab1, tab2, tab3 = st.tabs(["✍️ 手動新增題目", "📁 從 Word 匯入", "🚀 選題與匯出"])

# === Tab 1: 手動輸入 ===
with tab1:
    st.subheader("新增單一題目")
    
    # 第一列：分類設定
    col_cat1, col_cat2, col_cat3 = st.columns(3)
    with col_cat1:
        new_q_source = st.selectbox("來源", SOURCES)
    with col_cat2:
        # 章節選單
        chap_list = list(PHYSICS_CHAPTERS.keys())
        new_q_chap = st.selectbox("章節", chap_list)
    with col_cat3:
        # 根據章節動態產生單元選單
        unit_list = PHYSICS_CHAPTERS[new_q_chap]
        new_q_unit = st.selectbox("單元", unit_list)

    # 第二列：題型與答案
    c1, c2 = st.columns([1, 3])
    with c1:
        new_q_type = st.selectbox("題型", ["Single", "Multi", "Fill"], format_func=lambda x: {'Single':'單選題', 'Multi':'多選題', 'Fill':'填充題'}[x])
    with c2:
        new_q_ans = st.text_input("正確答案", placeholder="選擇題填代號(如 A)，填充題填文字")

    new_q_content = st.text_area("題目內容", height=100, placeholder="請輸入題目敘述...")
    new_q_image = st.file_uploader("上傳圖片 (選用)", type=['png', 'jpg', 'jpeg'])
    
    new_q_options = []
    if new_q_type in ["Single", "Multi"]:
        opts_text = st.text_area("選項 (每一行一個選項)", height=150, placeholder="選項 A\n選項 B\n選項 C\n選項 D")
        if opts_text:
            new_q_options = [line.strip() for line in opts_text.split('\n') if line.strip()]

    if st.button("➕ 加入題庫", type="secondary"):
        if not new_q_content:
            st.error("請輸入題目內容")
        elif new_q_type != 'Fill' and not new_q_options:
            st.error("選擇題必須提供選項")
        else:
            q_id = len(st.session_state['question_pool']) + 1
            img_bytes = new_q_image.getvalue() if new_q_image else None

            new_q = Question(
                new_q_type, new_q_content, new_q_options, new_q_ans, q_id, 
                image_data=img_bytes, 
                source=new_q_source, 
                chapter=new_q_chap, 
                unit=new_q_unit
            )
            st.session_state['question_pool'].append(new_q)
            st.success(f"已加入題目！分類：{new_q_source} / {new_q_unit}")

# === Tab 2: Word 匯入 ===
with tab2:
    st.subheader("批次匯入題目")
    st.write("支援標籤：`[Src:來源]`, `[Chap:章節]`, `[Unit:單元]`。")
    uploaded_file = st.file_uploader("上傳 Word (.docx) 檔案", type=['docx'])
    
    if uploaded_file:
        if st.button("解析並加入題庫"):
            try:
                imported_qs = parse_docx(uploaded_file.read())
                if imported_qs:
                    st.session_state['question_pool'].extend(imported_qs)
                    st.success(f"成功匯入 {len(imported_qs)} 題！")
                else:
                    st.warning("未偵測到題目。")
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
                # 顯示詳細分類標籤
                tags = f"[{q.source}] {q.unit}"
                with st.expander(f"{i+1}. {tags} {type_badge} {q.content.splitlines()[0][:30]}..."):
                    st.caption(f"完整分類：{q.chapter} > {q.unit}")
                    st.markdown(f"**題目**：\n{q.content}")
                    if q.image_data:
                        st.image(q.image_data, caption="題目附圖", width=300)
                    if q.options:
                        for idx, opt in enumerate(q.options):
                            st.text(f"({chr(65+idx)}) {opt}")
                    st.markdown(f"**答案**：`{q.answer}`")
                    if st.button("🗑️ 刪除", key=f"del_{i}"):
                        st.session_state['question_pool'].pop(i)
                        st.rerun()

        st.divider()
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
