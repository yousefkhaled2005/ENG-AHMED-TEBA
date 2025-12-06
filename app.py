import streamlit as st
import pandas as pd
import xlrd
import openpyxl
import random
import io
import zipfile
import gc  # Garbage Collector for memory management
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter
import time

# =========================================================
#  🎨 إعدادات التصميم (UI/UX)
# =========================================================
st.set_page_config(
    page_title="مولد الاختبارات المتقدم (نسخة السحابة)",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    * {font-family: 'Cairo', sans-serif;}
    .main {direction: rtl;}
    .stAlert {direction: rtl; text-align: right;}
    div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 15px;
        border: 1px solid #e0e0e0;
    }
    .log-container {
        background-color: #1e1e1e;
        color: #00ff00;
        font-family: 'Courier New', monospace;
        padding: 10px;
        border-radius: 5px;
        height: 200px;
        overflow-y: auto;
        font-size: 12px;
        direction: ltr;
        text-align: left;
    }
    h1, h2, h3, h4, p, label {text-align: right;}
    .stButton button {width: 100%; font-weight: bold; font-size: 18px; padding: 10px;}
</style>
""", unsafe_allow_html=True)

# =========================================================
#  🧠 منطق الألوان المتقدم (High Accuracy Color Logic)
# =========================================================

def is_cell_highlighted(cell):
    """
    دالة ذكية تكتشف ما إذا كانت الخلية ملونة بأي لون 
    (غير الأبيض أو الشفاف) بدقة عالية.
    """
    if not cell.fill or not cell.fill.start_color:
        return False
    
    color = cell.fill.start_color
    
    # 1. حالة الألوان RGB Hex
    if color.type == 'rgb':
        # ألوان شفافة أو بيضاء تماماً
        if color.rgb in [None, '00000000', 'FFFFFFFF']: 
            return False
        # أي كود لون آخر يعتبر تظليل
        return True
        
    # 2. حالة Theme Colors (ألوان القوالب في إكسيل)
    elif color.type == 'theme':
        # عادة الثيمات تكون ملونة إلا إذا كانت الخلفية الافتراضية
        # هنا نفترض أن أي ثيم محدد هو لون مقصود
        return True
        
    # 3. حالات Index (قديم)
    elif color.type == 'indexed':
        if color.indexed != 64: # 64 عادة هو التلقائي/الشفاف
            return True
            
    return False

# =========================================================
#  🔧 دوال المعالجة (Core Logic)
# =========================================================

def normalize_text(text):
    if text is None: return ""
    text = str(text).strip()
    return text.replace('ة', 'ه').replace('ى', 'ي').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')

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

# --- File Readers ---

def _read_xls_questions(file_obj):
    # Old XLS Format (xls)
    try:
        content = file_obj.read()
        book = xlrd.open_workbook(file_contents=content, formatting_info=True)
    except: return pd.DataFrame()
    
    sh = None
    for n in book.sheet_names():
        if 'بنك' in n or 'اسئله' in normalize_text(n): sh = book.sheet_by_name(n); break
    if not sh: sh = book.sheet_by_index(0)
    
    hr = -1; cols = {'u': -1, 'q': -1, 'obj': -1}
    for r in range(min(20, sh.nrows)):
        vs = [normalize_text(sh.cell_value(r, c)) for c in range(sh.ncols)]
        if any('سؤال' in x for x in vs) and any('وحده' in x for x in vs):
            hr = r
            for c, v in enumerate(vs):
                if 'وحده' in v: cols['u'] = c
                elif 'سؤال' in v: cols['q'] = c
                elif 'هدف' in v: cols['obj'] = c
            break
    if hr == -1: return pd.DataFrame()
    
    start_opt = min(cols['u'], cols['q']) + 1; end_opt = max(cols['u'], cols['q'])
    opt_cols = list(range(start_opt, end_opt))
    data = []
    
    for r in range(hr+1, sh.nrows):
        u = normalize_text(sh.cell_value(r, cols['u']))
        q = str(sh.cell_value(r, cols['q'])).strip()
        cat = u
        if cols['obj'] != -1: cat += str(sh.cell_value(r, cols['obj']))
        if not u or not q: continue
        o_txt = []; o_clr = []
        for ci in opt_cols:
            if ci < sh.ncols:
                val = str(sh.cell_value(r, ci)).strip()
                # Reading Color Index for XLS
                clr = book.xf_list[sh.cell_xf_index(r, ci)].background.pattern_colour_index
                o_txt.append(val); o_clr.append(clr)
        
        corr = ""; v_idx = [i for i,x in enumerate(o_txt) if x]
        
        # Logic: Find unique color or any color not equal to 64 (Automatic)
        ci = -1
        for i in v_idx:
            # 64 is usually 'Automatic/No Color' in xls
            if o_clr[i] != 64: 
                ci = i; break
                
        if ci != -1: corr = o_txt[ci]
        real_opts = [x for x in o_txt if x]
        if real_opts and corr:
            data.append({'category': cat, 'unit':u, 'question':q, 'options':real_opts[:4], 'correct_text':corr})
    
    # Clean memory
    del book
    return pd.DataFrame(data)

def _read_xlsx_questions(file_obj):
    # Modern Excel Format (xlsx)
    try: 
        wb = openpyxl.load_workbook(file_obj, data_only=False) # Keep formatting
    except: return pd.DataFrame()
    
    sh = wb.active
    rows = list(sh.iter_rows())
    hr = -1; cols = {'u': -1, 'q': -1, 'obj': -1}
    
    # Header detection
    for r_idx, row in enumerate(rows[:20]):
        vs = [normalize_text(cell.value) for cell in row]
        if any('سؤال' in x for x in vs) and any('وحده' in x for x in vs):
            hr = r_idx
            for c_idx, v in enumerate(vs):
                if 'وحده' in v: cols['u'] = c_idx
                elif 'سؤال' in v: cols['q'] = c_idx
                elif 'هدف' in v: cols['obj'] = c_idx
            break
    if hr == -1: return pd.DataFrame()

    start_opt = min(cols['u'], cols['q']) + 1; end_opt = max(cols['u'], cols['q'])
    opt_cols = list(range(start_opt, end_opt))
    data = []

    for row in rows[hr+1:]:
        try:
            u = normalize_text(row[cols['u']].value)
            q = str(row[cols['q']].value if row[cols['q']].value else "").strip()
        except IndexError: continue
        
        cat = u
        if cols['obj'] != -1 and cols['obj'] < len(row): cat += str(row[cols['obj']].value)
        if not u or not q: continue

        o_txt = []; o_is_colored = []
        for ci in opt_cols:
            if ci < len(row):
                cell = row[ci]
                val = str(cell.value if cell.value else "").strip()
                # === NEW: Use the robust color checker ===
                is_colored = is_cell_highlighted(cell)
                o_txt.append(val); o_is_colored.append(is_colored)
        
        corr = ""
        valid_indices = [i for i, txt in enumerate(o_txt) if txt]
        
        # Find the first highlighted option
        for i in valid_indices:
            if o_is_colored[i]: 
                corr = o_txt[i]
                break
                
        real_opts = [x for x in o_txt if x]
        if real_opts and corr:
            data.append({'category': cat, 'unit':u, 'question':q, 'options':real_opts[:4], 'correct_text':corr})
            
    # Clean Memory immediately
    del wb
    del rows
    return pd.DataFrame(data)

def fetch_smart_questions(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.lower().endswith('.xls'):
        return _read_xls_questions(uploaded_file)
    elif uploaded_file.name.lower().endswith('.xlsx'):
        return _read_xlsx_questions(uploaded_file)
    return pd.DataFrame()

def get_master_pattern_from_file(uploaded_file, limit=30):
    if uploaded_file is None:
        return [random.randint(0,3) for _ in range(limit)]
    try:
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        sheet = wb.active
        pattern = []
        for row in sheet.iter_rows():
            found_in_row = -1
            for cell in row:
                if is_cell_highlighted(cell) and cell.value:
                    txt = str(cell.value).strip()
                    if 'أ' in txt or 'ا' in txt: found_in_row = 0
                    elif 'ب' in txt: found_in_row = 1
                    elif 'ج' in txt: found_in_row = 2
                    elif 'د' in txt: found_in_row = 3
                    if found_in_row != -1: pattern.append(found_in_row); break
            if found_in_row == -1: # Random fallback if row is confusing
                pattern.append(random.randint(0,3))
                
        while len(pattern) < limit: pattern.append(random.randint(0,3))
        del wb # Clean memory
        return pattern
    except:
        return [random.randint(0,3) for _ in range(limit)]

def generate_balanced_exam(all_data_df, total):
    if all_data_df.empty: return pd.DataFrame()
    grp = all_data_df.groupby('category'); cats = list(grp.groups.keys())
    if not cats: return pd.DataFrame()
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
#  🖥️ واجهة المستخدم
# =========================================================

st.title("📄 مولد الاختبارات المتقدم (نسخة السحابة)")

col_log, col_banks = st.columns([1, 2])
with col_log:
    st.subheader(">_ سجل العمليات")
    log_placeholder = st.empty()
    if 'logs' not in st.session_state: st.session_state.logs = ["System Ready..."]
    
    def log(msg):
        st.session_state.logs.append(f"> {msg}")
        # Keep only last 15 logs to save memory/rendering
        if len(st.session_state.logs) > 15: st.session_state.logs.pop(0)
        log_txt = "\n".join(st.session_state.logs)
        log_placeholder.markdown(f'<div class="log-container"><pre>{log_txt}</pre></div>', unsafe_allow_html=True)
        time.sleep(0.05)

with col_banks:
    st.subheader("1️⃣ بنوك الأسئلة")
    uploaded_banks = st.file_uploader(
        "يمكنك رفع ملفات XLS أو XLSX (سيقرأ أي لون كإجابة)", 
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
    with st.container():
        c1, c2 = st.columns([8, 1])
        with c1: st.markdown(f"#### 📂 {model['name']}")
        with c2: st.markdown(f"_{model['folder']}_")
        col_temp, col_key = st.columns(2)
        with col_temp:
            t_file = st.file_uploader("التمبلت (Word)", type=['docx'], key=f"t_{model['key']}")
        with col_key:
            k_file = st.file_uploader("مفتاح الإجابة (Excel)", type=['xlsx'], key=f"k_{model['key']}")
        model_inputs[model['name']] = {"folder": model['folder'], "template": t_file, "key": k_file}

st.markdown("---")
max_q_count = st.number_input("عدد الأسئلة", min_value=1, value=30)

if st.button("🚀 بدء المعالجة", use_container_width=True):
    if not uploaded_banks:
        st.error("⚠️ يرجى رفع بنك أسئلة!")
    else:
        try:
            log("بدء المعالجة...")
            all_dfs = []
            
            # معالجة البنوك مع تنظيف الذاكرة
            for f in uploaded_banks:
                log(f"قراءة: {f.name}")
                df = fetch_smart_questions(f)
                if not df.empty: all_dfs.append(df)
                gc.collect() # Force memory cleanup
            
            if not all_dfs:
                st.error("لم يتم العثور على أسئلة!")
            else:
                BIG_DF = pd.concat(all_dfs).reset_index(drop=True)
                log(f"إجمالي الأسئلة: {len(BIG_DF)}")
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for m_name, inputs in model_inputs.items():
                        log(f"بناء نموذج: {m_name}")
                        
                        # Get Pattern
                        pattern = get_master_pattern_from_file(inputs['key'], limit=max_q_count)
                        gc.collect()

                        # Generate Exam Data
                        exam_df = generate_balanced_exam(BIG_DF, max_q_count)
                        if len(exam_df) > max_q_count: exam_df = exam_df.iloc[:max_q_count]
                        exam_df = exam_df.reset_index(drop=True)

                        # Create Word Doc
                        if inputs['template']:
                            inputs['template'].seek(0)
                            doc = Document(inputs['template'])
                        else:
                            doc = Document()
                        
                        set_section_rtl_and_margins(doc.sections[0])
                        # Add header if new doc (skipped if template used mostly)
                        if not inputs['template']:
                            doc.add_paragraph("")
                            sub = doc.add_paragraph(); fix_paragraph_alignment(sub)
                            run_st = sub.add_run("اختر الإجابة الصحيحة :")
                            force_font(run_st, size=14, is_bold=True)
                        
                        for idx, row in enumerate(exam_df.to_dict('records')):
                            target_idx = pattern[idx % len(pattern)]
                            final_opts = force_align_options(row['options'], row['correct_text'], target_idx)
                            add_question_block(doc, idx+1, row['question'], final_opts)
                        
                        # Save to Zip
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_path = f"{inputs['folder']}/{m_name}.docx"
                        zf.writestr(zip_path, doc_io.getvalue())
                        
                        # Memory cleanup per iteration
                        del doc
                        del doc_io
                        gc.collect()
                
                log("✅ تم الانتهاء بنجاح!")
                st.success("تم تجهيز الملفات!")
                st.download_button(
                    label="📥 تحميل النماذج (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Final_Exams_Optimized.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"حدث خطأ: {e}")
            log(f"Error: {str(e)}")
