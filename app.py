import streamlit as st
import pandas as pd
import xlrd
import openpyxl
import random
import io
import zipfile
import gc
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter
import time

# =========================================================
# 🎨 واجهة المستخدم (التصميم)
# =========================================================
st.set_page_config(
    page_title="منصة المهندس أحمد - جامعة طيبة",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;900&display=swap');
    * {font-family: 'Tajawal', sans-serif;}
    .main {background-color: #f8f9fa;}
    .header-container {
        background: linear-gradient(90deg, #004d40 0%, #00695c 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .header-title {font-size: 32px; font-weight: 900; margin: 0;}
    .header-subtitle {font-size: 18px; opacity: 0.9; margin-top: 5px;}
    div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border: 1px solid #eaeaea;
    }
    .log-container {
        background-color: #1e1e1e;
        color: #00e676;
        font-family: 'Courier New', monospace;
        padding: 15px;
        border-radius: 8px;
        height: 250px;
        overflow-y: auto;
        font-size: 13px;
        direction: ltr;
        text-align: left;
        border: 2px solid #333;
    }
    .stButton button {
        width: 100%;
        background-color: #00695c;
        color: white;
        font-weight: bold;
        font-size: 18px;
        padding: 12px;
        border-radius: 8px;
        border: none;
        transition: all 0.3s;
    }
    .stButton button:hover {background-color: #004d40; box-shadow: 0 4px 10px rgba(0,0,0,0.2);}
    h1, h2, h3, h4, p, label {text-align: right; direction: rtl;}
    .stFileUploader label {font-weight: bold; color: #333;}
</style>
""", unsafe_allow_html=True)

# =========================================================
# 🔧 دوال المعالجة (وضع الطوارئ - Brute Force Mode)
# =========================================================

def normalize_text(text):
    if text is None: return ""
    text = str(text).strip()
    return text.replace('ة', 'ه').replace('ى', 'ي').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ؤ', 'و').replace('ئ', 'ي')

def clean_for_comp(text): return normalize_text(text).replace(" ", "")

def force_align_options(options, correct_text, target_idx):
    final_opts = options[:]
    while len(final_opts) < 4: final_opts.append("---")
    current_idx = -1
    clean_corr = clean_for_comp(correct_text)
    for i, opt in enumerate(final_opts):
        if clean_for_comp(str(opt)) == clean_corr: current_idx = i; break
    if current_idx != -1:
        if current_idx != target_idx:
            temp = final_opts[target_idx]; final_opts[target_idx] = final_opts[current_idx]; final_opts[current_idx] = temp
    else: final_opts[target_idx] = correct_text
    return final_opts

# --- Word Formatting ---
def set_section_rtl_and_margins(section):
    sectPr = section._sectPr
    if not sectPr.find(qn('w:bidi')):
        bidi = OxmlElement('w:bidi'); bidi.set(qn('w:val'), '1'); sectPr.append(bidi)
    section.left_margin = Inches(0.4); section.right_margin = Inches(0.4)
    section.top_margin = Inches(0.5); section.bottom_margin = Inches(0.5)

def fix_paragraph_alignment(paragraph):
    p_fmt = paragraph.paragraph_format
    p_fmt.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p_fmt.bidi = True
    p_fmt.left_indent = Pt(0); p_fmt.right_indent = Pt(0); p_fmt.first_line_indent = Pt(0)
    p_fmt.space_before = Pt(4); p_fmt.space_after = Pt(4)

def configure_table_layout(table):
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None: tbl_pr = table._tbl.add_tblPr()
    bidi = OxmlElement('w:bidiVisual'); tbl_pr.append(bidi)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'right'); tbl_pr.append(jc)
    tbl_borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}'); border.set(qn('w:val'), 'nil'); tbl_borders.append(border)
    tbl_pr.append(tbl_borders)
    tbl_cell_mar = OxmlElement('w:tblCellMar')
    for m in ['top', 'start', 'bottom', 'end']:
        width = OxmlElement(f'w:{m}'); width.set(qn('w:w'), '0'); width.set(qn('w:type'), 'dxa'); tbl_cell_mar.append(width)
    tbl_pr.append(tbl_cell_mar)

def force_font(run, size=14, is_bold=False):
    run.font.name = 'Times New Roman'; run.font.size = Pt(size); run.bold = is_bold
    rPr = run._element.get_or_add_rPr()
    fonts = OxmlElement('w:rFonts'); fonts.set(qn('w:ascii'), 'Times New Roman'); fonts.set(qn('w:hAnsi'), 'Times New Roman'); fonts.set(qn('w:eastAsia'), 'Times New Roman'); fonts.set(qn('w:cs'), 'Times New Roman'); rPr.append(fonts)
    sz = OxmlElement('w:szCs'); sz.set(qn('w:val'), str(int(size * 2))); rPr.append(sz)
    b = OxmlElement('w:bCs'); b.set(qn('w:val'), '1' if is_bold else '0'); rPr.append(b)
    if not rPr.find(qn('w:rtl')): rPr.append(OxmlElement('w:rtl'))

def add_question_block(doc, q_num, q_text, options):
    table = doc.add_table(rows=2, cols=4)
    configure_table_layout(table)
    q_cell = table.cell(0, 0).merge(table.cell(0, 3))
    q_p = q_cell.paragraphs[0]
    q_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; q_p.paragraph_format.bidi = True
    run = q_p.add_run(f"{q_num}) {q_text}"); force_font(run, size=14, is_bold=True)
    chars = ['أ', 'ب', 'ج', 'د']
    for i, opt_text in enumerate(options):
        opt_cell = table.cell(1, i); opt_p = opt_cell.paragraphs[0]
        opt_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; opt_p.paragraph_format.bidi = True
        lbl = chars[i] if i < len(chars) else "-"
        run_opt = opt_p.add_run(f"{lbl}. {opt_text}"); force_font(run_opt, size=14, is_bold=False)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# --- منطق كشف الألوان المطور ---
def is_cell_colored(cell):
    if not cell.fill or not cell.fill.start_color: return False
    c = cell.fill.start_color
    if c.type == 'rgb':
        if str(c.rgb).upper() in ['00000000', 'FFFFFFFF', 'NONE']: return False
        return True 
    if c.type == 'theme':
        if c.theme in [0, 1] and (c.tint == 0.0): return False
        return True
    if c.type == 'indexed':
        if c.indexed == 64: return False
        return True
    return False

def _read_xlsx_questions_brute_force(file_obj):
    """
    دالة تقرأ الملف "بالعافية" بدون الاعتماد على أسماء الأعمدة الدقيقة
    """
    try: wb = openpyxl.load_workbook(file_obj, data_only=False)
    except: return pd.DataFrame()
    
    sh = wb.active
    rows = list(sh.iter_rows())
    
    # 1. البحث عن صف العناوين (تقريبي)
    header_row_idx = -1
    q_col_idx = -1
    u_col_idx = -1
    
    # ابحث في أول 50 صف عن كلمة تشبه "سؤال"
    for r_idx, row in enumerate(rows[:50]):
        vs = [normalize_text(cell.value) for cell in row]
        for c_idx, val in enumerate(vs):
            if 'سؤال' in val or 'سوال' in val or 'الاسيله' in val:
                header_row_idx = r_idx
                q_col_idx = c_idx
            if 'وحده' in val or 'وحدة' in val:
                u_col_idx = c_idx
        if q_col_idx != -1: break
    
    # إذا لم يجد هيدر، افترض أن آخر عمود فيه كلام هو السؤال!
    if q_col_idx == -1:
        header_row_idx = 0 # ابدأ من الأول وخلاص
        # ابحث عن العمود اللي فيه أطول نصوص (غالباً هو الأسئلة)
        max_len = 0
        best_col = -1
        for r in rows[1:min(20, len(rows))]:
            for c_idx, cell in enumerate(r):
                if cell.value and len(str(cell.value)) > max_len:
                    max_len = len(str(cell.value))
                    best_col = c_idx
        q_col_idx = best_col

    # إذا لسه مافيش، استسلم
    if q_col_idx == -1: return pd.DataFrame()

    data = []
    
    # ابدأ القراءة من بعد صف الهيدر
    start_row = header_row_idx + 1 if header_row_idx != -1 else 0
    
    for row in rows[start_row:]:
        try:
            # استخراج السؤال
            q_val = row[q_col_idx].value
            if not q_val: continue
            q = str(q_val).strip()
            if len(q) < 3: continue # تجاهل الأرقام أو الرموز
            
            # استخراج الوحدة (أو وضع افتراضي)
            u = "عام"
            if u_col_idx != -1:
                try: u = str(row[u_col_idx].value).strip() if row[u_col_idx].value else "عام"
            
                except: pass
            
            # استخراج الاختيارات:
            # نأخذ كل الخلايا في الصف، ما عدا خلية السؤال والوحدة
            o_txt = []
            corr = ""
            
            for c_idx, cell in enumerate(row):
                if c_idx == q_col_idx or c_idx == u_col_idx: continue # تخطي السؤال والوحدة
                
                val = str(cell.value if cell.value else "").strip()
                if val:
                    # تصفية الكلمات التي لا يمكن أن تكون خيارات (مثل العناوين الجانبية)
                    if val in ["سهل", "صعب", "متوسط", "المستوى", "الهدف"]: continue
                    
                    o_txt.append(val)
                    if is_cell_colored(cell):
                        corr = val
            
            # ترتيب الخيارات والحل
            real_opts = o_txt
            
            # Fallback 1: إذا لم يجد لون، خذ أول خيار واعتبره هو الصح (مؤقتاً)
            if real_opts and not corr:
                corr = real_opts[0] # Force answer
            
            if real_opts and corr and len(real_opts) >= 2:
                data.append({'category': u, 'unit':u, 'question':q, 'options':real_opts[:4], 'correct_text':corr})
                
        except Exception: continue

    return pd.DataFrame(data)

def _read_xls_questions_simple(file_obj):
    # XLS fallback simpler version
    try:
        content = file_obj.read()
        book = xlrd.open_workbook(file_contents=content, formatting_info=True)
    except: return pd.DataFrame()
    sh = book.sheet_by_index(0)
    
    # Assume last column is Question, first is Unit
    q_col = sh.ncols - 1
    u_col = 0
    # Search for header to confirm
    for r in range(min(20, sh.nrows)):
        for c in range(sh.ncols):
            v = normalize_text(sh.cell_value(r, c))
            if 'سؤال' in v or 'سوال' in v: q_col = c
            if 'وحده' in v: u_col = c
            
    data = []
    for r in range(sh.nrows):
        q = str(sh.cell_value(r, q_col)).strip()
        if not q or len(q)<3: continue
        u = str(sh.cell_value(r, u_col)).strip() if u_col != q_col else "عام"
        
        o_txt = []
        corr = ""
        for c in range(sh.ncols):
            if c == q_col or c == u_col: continue
            val = str(sh.cell_value(r, c)).strip()
            if val:
                o_txt.append(val)
                # Check color
                xf = book.xf_list[sh.cell_xf_index(r, c)]
                bg = xf.background.pattern_colour_index
                if bg != 64: corr = val
        
        if o_txt and not corr: corr = o_txt[0] # Force answer
        
        if o_txt and corr:
             data.append({'category': u, 'unit':u, 'question':q, 'options':o_txt[:4], 'correct_text':corr})
             
    return pd.DataFrame(data)

def fetch_smart_questions(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.lower().endswith('.xls'): return _read_xls_questions_simple(uploaded_file)
    elif uploaded_file.name.lower().endswith('.xlsx'): return _read_xlsx_questions_brute_force(uploaded_file)
    return pd.DataFrame()

def get_master_pattern_from_file(uploaded_file, limit=30):
    # Simplified pattern reader
    if uploaded_file is None: return [random.randint(0,3) for _ in range(limit)]
    try:
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        sheet = wb.active
        pattern = []
        for row in sheet.iter_rows():
            found = False
            for i, cell in enumerate(row):
                if is_cell_colored(cell) and cell.value:
                    txt = str(cell.value).strip()
                    if 'أ' in txt or 'ا' in txt: pattern.append(0)
                    elif 'ب' in txt: pattern.append(1)
                    elif 'ج' in txt: pattern.append(2)
                    elif 'د' in txt: pattern.append(3)
                    else: pattern.append(random.randint(0,3))
                    found = True; break
            if not found and len(pattern) < limit: pattern.append(random.randint(0,3))
            if len(pattern) >= limit: break
        while len(pattern) < limit: pattern.append(random.randint(0,3))
        return pattern
    except: return [random.randint(0,3) for _ in range(limit)]

def generate_balanced_exam(all_data_df, total):
    if all_data_df.empty: return pd.DataFrame()
    # If category/unit is missing or mostly "عام", just random sample
    if len(all_data_df['category'].unique()) < 2:
         return all_data_df.sample(n=min(total, len(all_data_df))).reset_index(drop=True)
         
    grp = all_data_df.groupby('category'); cats = list(grp.groups.keys())
    base = total // len(cats); rem = total % len(cats)
    sel = []; random.shuffle(cats)
    for i, cat in enumerate(cats):
        q = base + (1 if i < rem else 0)
        df = grp.get_group(cat)
        sel.append(df.sample(n=q) if len(df)>=q else df)
    res = pd.concat(sel)
    if len(res) < total:
        left = all_data_df[~all_data_df.index.isin(res.index)]
        need = total - len(res)
        if not left.empty: res = pd.concat([res, left.sample(n=need) if len(left)>=need else left])
    return res.sample(frac=1).reset_index(drop=True)

# =========================================================
# 🖥️ لوحة التحكم
# =========================================================

st.markdown("""
<div class="header-container">
    <div class="header-title">منصة المهندس أحمد</div>
    <div class="header-subtitle">جامعة طيبة - كلية الهندسة</div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=100)
    st.title("لوحة التحكم")
    st.markdown("---")
    st.subheader(">_ سجل العمليات")
    log_placeholder = st.empty()
    if 'logs' not in st.session_state: st.session_state.logs = ["System Ready..."]
    
    def log(msg):
        st.session_state.logs.append(f"> {msg}")
        if len(st.session_state.logs) > 15: st.session_state.logs.pop(0)
        log_txt = "\n".join(st.session_state.logs)
        log_placeholder.markdown(f'<div class="log-container"><pre>{log_txt}</pre></div>', unsafe_allow_html=True)
        time.sleep(0.05)
    
    st.info("الوضع: قراءة قسرية (Brute Force Mode) - سيتم استخراج الأسئلة بأي ثمن.")

st.subheader("1️⃣ رفع بنوك الأسئلة")
uploaded_banks = st.file_uploader(
    "ارفع ملفات البنك (xls أو xlsx)", 
    accept_multiple_files=True, 
    type=['xls', 'xlsx'],
    key="banks_uploader"
)

st.markdown("---")

st.subheader("2️⃣ إعدادات النماذج")
models_config = [
    {"name": "ا_صباحي", "folder": "صباحي", "key": "am1"},
    {"name": "ب_صباحي", "folder": "صباحي", "key": "am2"},
    {"name": "ا_مسائي", "folder": "مسائي", "key": "pm1"},
    {"name": "ب_مسائي", "folder": "مسائي", "key": "pm2"},
    {"name": "ا_دور_ثاني", "folder": "دور_ثاني", "key": "sec1"},
    {"name": "ب_دور_ثاني", "folder": "دور_ثاني", "key": "sec2"},
]

model_inputs = {}
for model in models_config:
    with st.expander(f"📌 إعدادات نموذج: {model['name']}", expanded=False):
        c1, c2 = st.columns(2)
        with c1: t_file = st.file_uploader("التمبلت (Word)", type=['docx'], key=f"t_{model['key']}")
        with c2: k_file = st.file_uploader("مفتاح الإجابة (Excel)", type=['xlsx'], key=f"k_{model['key']}")
        model_inputs[model['name']] = {"folder": model['folder'], "template": t_file, "key": k_file}

st.markdown("---")
c_num, c_btn = st.columns([1, 3])
with c_num:
    max_q_count = st.number_input("عدد الأسئلة لكل نموذج", min_value=1, value=30)
with c_btn:
    st.text("") 
    st.text("") 
    start_btn = st.button("🚀 بدء توليد الاختبارات", use_container_width=True)

if start_btn:
    gc.collect()
    if not uploaded_banks:
        st.error("⚠️ خطأ: يجب رفع بنك أسئلة واحد على الأقل.")
    else:
        try:
            log("بدء المعالجة...")
            progress_bar = st.progress(0)
            all_dfs = []
            
            total_files = len(uploaded_banks)
            for i, f in enumerate(uploaded_banks):
                log(f"جاري قراءة الملف: {f.name}")
                df = fetch_smart_questions(f)
                if not df.empty: 
                    all_dfs.append(df)
                    log(f"✅ تم قراءة {len(df)} سؤال بالقوة.")
                else: 
                    log(f"⚠️ فشل حتى القراءة القسرية في {f.name}.")
                progress_bar.progress((i + 1) / (total_files * 2)) 
                gc.collect()
            
            if not all_dfs:
                st.error("❌ لم يتم العثور على بيانات حتى مع الوضع القسري. الملف قد يكون مشفراً أو فارغاً تماماً.")
            else:
                BIG_DF = pd.concat(all_dfs).reset_index(drop=True)
                log(f"📊 المجموع: {len(BIG_DF)} سؤال.")
                del all_dfs; gc.collect()
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    total_models = len(model_inputs)
                    for idx, (m_name, inputs) in enumerate(model_inputs.items()):
                        log(f"جاري بناء: {m_name}")
                        pattern = get_master_pattern_from_file(inputs['key'], limit=max_q_count)
                        gc.collect()

                        exam_df = generate_balanced_exam(BIG_DF, max_q_count)
                        if len(exam_df) > max_q_count: exam_df = exam_df.iloc[:max_q_count]
                        exam_df = exam_df.reset_index(drop=True)

                        if inputs['template']:
                            inputs['template'].seek(0); doc = Document(inputs['template'])
                        else: doc = Document()
                        
                        set_section_rtl_and_margins(doc.sections[0])
                        if not inputs['template']:
                            doc.add_paragraph("")
                            sub = doc.add_paragraph(); fix_paragraph_alignment(sub)
                            run_st = sub.add_run("اختر الإجابة الصحيحة :")
                            force_font(run_st, size=14, is_bold=True)
                        
                        for i, row in enumerate(exam_df.to_dict('records')):
                            target_idx = pattern[i % len(pattern)]
                            final_opts = force_align_options(row['options'], row['correct_text'], target_idx)
                            add_question_block(doc, i+1, row['question'], final_opts)
                        
                        doc_io = io.BytesIO(); doc.save(doc_io)
                        zip_path = f"{inputs['folder']}/{m_name}.docx"
                        zf.writestr(zip_path, doc_io.getvalue())
                        
                        del doc; del doc_io; del exam_df; gc.collect()
                        progress_bar.progress(0.5 + ((idx + 1) / (total_models * 2)))
                
                progress_bar.progress(100)
                log("✅ تمت العملية!")
                st.balloons()
                st.success(f"تم إنشاء {len(BIG_DF)} سؤال. (ملاحظة: إذا لم تكن الإجابات ملونة، تم اختيار الخيار الأول افتراضياً).")
                
                st.download_button(
                    label="📥 تحميل الملفات (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Taibah_Exams_ForceMode.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                del BIG_DF; del zip_buffer; gc.collect()
                
        except Exception as e:
            st.error(f"حدث خطأ: {e}")
            log(f"Error: {str(e)}")
