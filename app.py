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
# 🎨 تحسين الواجهة والتصميم
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
# 🔧 دوال المعالجة (Logic) - تم التحديث للتعامل مع ألوان الإكسل بدقة
# =========================================================

def normalize_text(text):
    if text is None: return ""
    text = str(text).strip()
    return text.replace('ة', 'ه').replace('ى', 'ي').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ؤ', 'و')

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

def get_xlsx_color_signature(cell):
    """
    استخراج بصمة اللون للخلية بغض النظر عن نوعه (Theme, RGB, Indexed)
    للمقارنة بين الخلايا في نفس الصف
    """
    if not cell.fill or not cell.fill.start_color:
        return "NONE"
    
    color = cell.fill.start_color
    if color.type == 'rgb':
        # ignore transparent or pure white if logic demands, but better keep distinct
        return f"RGB:{color.rgb}"
    elif color.type == 'theme':
        return f"THEME:{color.theme}:TINT:{color.tint}"
    elif color.type == 'indexed':
        return f"INDEX:{color.indexed}"
    return "NONE"

def _read_xlsx_questions(file_obj):
    try: wb = openpyxl.load_workbook(file_obj, data_only=False)
    except: return pd.DataFrame()
    
    sh = wb.active
    rows = list(sh.iter_rows())
    hr = -1; cols = {'u': -1, 'q': -1, 'obj': -1}
    
    # Header Detection
    for r_idx, row in enumerate(rows[:25]): # Scan first 25 rows
        vs = [normalize_text(cell.value) for cell in row]
        # Check for standard headers (Question, Unit)
        if any('سؤال' in x for x in vs) and any('وحده' in x for x in vs):
            hr = r_idx
            for c_idx, v in enumerate(vs):
                if 'وحده' in v: cols['u'] = c_idx
                elif 'سؤال' in v: cols['q'] = c_idx
                elif 'هدف' in v: cols['obj'] = c_idx
            break
            
    if hr == -1: return pd.DataFrame()

    # Determine Options Range (Between Unit and Question, usually)
    start_opt = min(cols['u'], cols['q']) + 1
    end_opt = max(cols['u'], cols['q'])
    opt_cols = list(range(start_opt, end_opt))
    data = []

    for row in rows[hr+1:]:
        try:
            u_val = row[cols['u']].value
            q_val = row[cols['q']].value
            if not u_val or not q_val: continue # Skip empty rows or sub-headers
            
            u = normalize_text(u_val)
            q = str(q_val).strip()
        except IndexError: continue
        
        cat = u
        if cols['obj'] != -1 and cols['obj'] < len(row): 
            if row[cols['obj']].value: cat += str(row[cols['obj']].value)

        o_txt = []
        o_sigs = [] # Color signatures
        
        for ci in opt_cols:
            if ci < len(row):
                cell = row[ci]
                val = str(cell.value if cell.value else "").strip()
                sig = get_xlsx_color_signature(cell)
                o_txt.append(val)
                o_sigs.append(sig)
        
        # --- Smart Color Logic ---
        # Find the "Unique" color in this row.
        # If we have [White, White, Green, White], Green is the answer.
        corr = ""
        valid_indices = [i for i, txt in enumerate(o_txt) if txt]
        valid_sigs = [o_sigs[i] for i in valid_indices]
        
        if valid_sigs:
            # Check if there is a color that appears exactly ONCE while others appear multiple times
            # Or if there is a color different from the "mode" (most common color)
            counts = Counter(valid_sigs)
            
            # Scenario 1: One distinct color, others are same
            # e.g., {White: 3, Green: 1} -> Green is answer
            unique_sig = next((k for k, v in counts.items() if v == 1), None)
            
            found_idx = -1
            if unique_sig and len(counts) > 1:
                # We found a unique color!
                for i in valid_indices:
                    if o_sigs[i] == unique_sig:
                        found_idx = i; break
            
            # Scenario 2: If simple unique fails, try just finding non-default
            # (Fallback if logic is tricky, but Unique is usually safest for MCQ)
            if found_idx == -1 and len(counts) > 1:
                 # Try to exclude common "White/Transparent" signatures if possible
                 # But sticking to "Minority Color" is usually correct.
                 # Let's check for "Any color that is NOT the most common one"
                 most_common = counts.most_common(1)[0][0]
                 for i in valid_indices:
                     if o_sigs[i] != most_common:
                         found_idx = i; break

            if found_idx != -1:
                corr = o_txt[found_idx]

        real_opts = [x for x in o_txt if x]
        if real_opts and corr:
            data.append({'category': cat, 'unit':u, 'question':q, 'options':real_opts[:4], 'correct_text':corr})
            
    return pd.DataFrame(data)

def _read_xls_questions(file_obj):
    # (Same as before, XLS handling is usually stable)
    try:
        content = file_obj.read()
        book = xlrd.open_workbook(file_contents=content, formatting_info=True)
    except: return pd.DataFrame()
    
    sh = None
    for n in book.sheet_names():
        if 'بنك' in n or 'اسئله' in normalize_text(n): sh = book.sheet_by_name(n); break
    if not sh: sh = book.sheet_by_index(0)
    
    hr = -1; cols = {'u': -1, 'q': -1, 'obj': -1}
    for r in range(min(25, sh.nrows)):
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
                clr = book.xf_list[sh.cell_xf_index(r, ci)].background.pattern_colour_index
                o_txt.append(val); o_clr.append(clr)
        
        corr = ""; v_idx = [i for i,x in enumerate(o_txt) if x]; v_clr = [o_clr[i] for i in v_idx]
        ci = -1
        if v_clr:
            cnt = Counter(v_clr); uniq = next((k for k,v in cnt.items() if v==1), None)
            if uniq: 
                for i in v_idx: 
                    if o_clr[i] == uniq: ci=i; break
            else: 
                for i in v_idx: 
                    if o_clr[i]!=64: ci=i; break # 64 is standard white/auto
        if ci != -1: corr = o_txt[ci]
        real_opts = [x for x in o_txt if x]
        if real_opts and corr:
            data.append({'category': cat, 'unit':u, 'question':q, 'options':real_opts[:4], 'correct_text':corr})
    return pd.DataFrame(data)

def fetch_smart_questions(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.lower().endswith('.xls'): return _read_xls_questions(uploaded_file)
    elif uploaded_file.name.lower().endswith('.xlsx'): return _read_xlsx_questions(uploaded_file)
    return pd.DataFrame()

def get_master_pattern_from_file(uploaded_file, limit=30):
    if uploaded_file is None: return [random.randint(0,3) for _ in range(limit)]
    try:
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        sheet = wb.active
        pattern = []
        for row in sheet.iter_rows():
            sigs = []
            vals = []
            for cell in row:
                sigs.append(get_xlsx_color_signature(cell))
                vals.append(cell.value)
            
            # Logic: Find unique color
            counts = Counter(sigs)
            uniq = next((k for k, v in counts.items() if v == 1), None)
            
            found_in_row = -1
            if uniq and len(counts) > 1:
                # Find which index (0=A, 1=B...) matches the unique color AND has text
                # Simple logic for pattern template: Just find the colored cell's text content
                # "أ" or "ب" ...
                for i, sig in enumerate(sigs):
                    if sig == uniq and vals[i]:
                        txt = str(vals[i]).strip()
                        if 'أ' in txt or 'ا' in txt: found_in_row = 0
                        elif 'ب' in txt: found_in_row = 1
                        elif 'ج' in txt: found_in_row = 2
                        elif 'د' in txt: found_in_row = 3
                        break
            
            if found_in_row != -1: pattern.append(found_in_row)
            else: pattern.append(random.randint(0,3))
            
            if len(pattern) >= limit: break
            
        while len(pattern) < limit: pattern.append(random.randint(0,3))
        return pattern
    except: return [random.randint(0,3) for _ in range(limit)]

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
# 🖥️ واجهة المستخدم
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
    
    st.info("نظام ذكي لتوليد الاختبارات المتوازنة من ملفات الإكسل.")

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
                if not df.empty: all_dfs.append(df)
                else: log(f"تحذير: الملف {f.name} لا يحتوي على أسئلة مقروءة! تأكد من التظليل.")
                progress_bar.progress((i + 1) / (total_files * 2)) 
                gc.collect()
            
            if not all_dfs:
                st.error("لم يتم العثور على أسئلة صالحة! (تأكد أن الإجابة الصحيحة مظللة بلون مختلف عن البقية)")
            else:
                BIG_DF = pd.concat(all_dfs).reset_index(drop=True)
                log(f"تم استخراج {len(BIG_DF)} سؤال بنجاح.")
                del all_dfs; gc.collect()
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    total_models = len(model_inputs)
                    for idx, (m_name, inputs) in enumerate(model_inputs.items()):
                        log(f"جاري بناء النموذج: {m_name}")
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
                log("✅ تمت العملية بنجاح!")
                st.balloons()
                st.success(f"تم إنشاء {len(BIG_DF)} سؤال في الملفات وجاهزة للتحميل.")
                
                st.download_button(
                    label="📥 تحميل الملفات المضغوطة (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Taibah_Exams_Final.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                del BIG_DF; del zip_buffer; gc.collect()
                
        except Exception as e:
            st.error(f"حدث خطأ غير متوقع: {e}")
            log(f"Error: {str(e)}")
