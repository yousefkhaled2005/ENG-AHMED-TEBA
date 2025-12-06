import streamlit as st
import pandas as pd
import xlrd
import openpyxl
import random
import io
import zipfile
import gc
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import Counter

# =========================================================
#  🎨 إعدادات الصفحة (Pro UI)
# =========================================================
st.set_page_config(
    page_title="نظام إدارة الاختبارات الذكي",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    * {font-family: 'Cairo', sans-serif;}
    .main {direction: rtl;}
    .stButton button {
        background-color: #2e86de; color: white; 
        font-size: 20px; padding: 15px; border-radius: 8px;
        transition: 0.3s;
    }
    .stButton button:hover {background-color: #1e3799; border: none;}
    .success-box {padding: 15px; background-color: #dff9fb; border-radius: 10px; color: #130f40; border: 1px solid #c7ecee;}
    .error-box {padding: 10px; background-color: #ff7979; border-radius: 5px; color: white;}
    div[data-testid="stExpander"] {border: 1px solid #ddd; border-radius: 8px;}
</style>
""", unsafe_allow_html=True)

# =========================================================
#  ⚙️ دوال المعالجة (Core Logic)
# =========================================================

def normalize_text(text):
    if pd.isna(text) or text == "": return ""
    text = str(text).strip()
    return text.replace('ة', 'ه').replace('ى', 'ي').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')

def clean_for_comp(text): return normalize_text(text).replace(" ", "")

def force_align_options(options, correct_text, target_idx):
    # نسخة آمنة: تتأكد من وجود بيانات قبل العمل
    final_opts = list(options) if options else []
    while len(final_opts) < 4: final_opts.append("---")
    
    current_idx = -1
    clean_corr = clean_for_comp(correct_text)
    
    for i, opt in enumerate(final_opts):
        if clean_for_comp(str(opt)) == clean_corr: current_idx = i; break
            
    if current_idx != -1:
        if current_idx != target_idx:
            try:
                final_opts[target_idx], final_opts[current_idx] = final_opts[current_idx], final_opts[target_idx]
            except IndexError: pass 
    else:
        # Fallback
        if 0 <= target_idx < 4:
            final_opts[target_idx] = correct_text
            
    return final_opts[:4]

# --- Word Helpers ---
def set_section_rtl(section):
    sectPr = section._sectPr
    if not sectPr.find(qn('w:bidi')):
        bidi = OxmlElement('w:bidi'); bidi.set(qn('w:val'), '1'); sectPr.append(bidi)
    section.left_margin = Inches(0.4); section.right_margin = Inches(0.4)
    section.top_margin = Inches(0.5); section.bottom_margin = Inches(0.5)

def fix_paragraph(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.bidi = True
    p.paragraph_format.left_indent = Pt(0); p.paragraph_format.right_indent = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(4)

def format_table(table):
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None: tbl_pr = table._tbl.add_tblPr()
    bidi = OxmlElement('w:bidiVisual'); tbl_pr.append(bidi)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'right'); tbl_pr.append(jc)
    tbl_borders = OxmlElement('w:tblBorders')
    for border in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{border}'); el.set(qn('w:val'), 'nil'); tbl_borders.append(el)
    tbl_pr.append(tbl_borders)
    tbl_cell_mar = OxmlElement('w:tblCellMar')
    for m in ['top', 'start', 'bottom', 'end']:
        w = OxmlElement(f'w:{m}'); w.set(qn('w:w'), '0'); w.set(qn('w:type'), 'dxa'); tbl_cell_mar.append(w)
    tbl_pr.append(tbl_cell_mar)

def set_font(run, size=14, bold=False):
    run.font.name = 'Times New Roman'; run.font.size = Pt(size); run.bold = bold
    rPr = run._element.get_or_add_rPr()
    fonts = OxmlElement('w:rFonts')
    fonts.set(qn('w:ascii'), 'Times New Roman'); fonts.set(qn('w:hAnsi'), 'Times New Roman')
    fonts.set(qn('w:eastAsia'), 'Times New Roman'); fonts.set(qn('w:cs'), 'Times New Roman')
    rPr.append(fonts)
    if not rPr.find(qn('w:rtl')): rPr.append(OxmlElement('w:rtl'))

def add_question(doc, num, text, opts):
    try:
        t = doc.add_table(rows=2, cols=4); format_table(t)
        # Question row (merged)
        c = t.cell(0, 0).merge(t.cell(0, 3))
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.bidi = True
        set_font(p.add_run(f"{num}) {text}"), 14, True)
        # Options row
        chars = ['أ', 'ب', 'ج', 'د']
        for i, opt in enumerate(opts):
            if i < 4:
                c2 = t.cell(1, i); p2 = c2.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.paragraph_format.bidi = True
                set_font(p2.add_run(f"{chars[i] if i<len(chars) else '-'}. {opt}"), 14, False)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
    except: pass

# --- Smart Readers (Memory Efficient) ---
# التخزين المؤقت (Caching) يمنع إعادة قراءة الملفات الثقيلة عند كل ضغطة زر
@st.cache_data(show_spinner=False)
def read_question_bank(file_bytes, filename):
    try:
        # Determine engine based on extension
        if filename.lower().endswith('.xls'):
            book = xlrd.open_workbook(file_contents=file_bytes.getvalue(), formatting_info=True)
            # ... (XLS Logic from previous code) ...
            sh = book.sheet_by_index(0)
            for n in book.sheet_names():
                if 'بنك' in n or 'اسئله' in normalize_text(n): sh = book.sheet_by_name(n); break
            
            # Logic extraction (Optimized)
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
            
            s_opt = min(cols['u'], cols['q']) + 1; e_opt = max(cols['u'], cols['q'])
            opt_cols = list(range(s_opt, e_opt))
            
            data = []
            for r in range(hr+1, sh.nrows):
                u = normalize_text(sh.cell_value(r, cols['u']))
                q = str(sh.cell_value(r, cols['q'])).strip()
                cat = u + (str(sh.cell_value(r, cols['obj'])) if cols['obj']!=-1 else "")
                if not u or not q: continue
                
                o_txt = []; o_clr = []
                for ci in opt_cols:
                    if ci < sh.ncols:
                        o_txt.append(str(sh.cell_value(r, ci)).strip())
                        o_clr.append(book.xf_list[sh.cell_xf_index(r, ci)].background.pattern_colour_index)
                
                corr = ""; v_idx = [i for i,x in enumerate(o_txt) if x]; v_clr = [o_clr[i] for i in v_idx]
                ci = -1
                if v_clr:
                    cnt = Counter(v_clr); uniq = next((k for k,v in cnt.items() if v==1), None)
                    if uniq: 
                        for i in v_idx: 
                            if o_clr[i] == uniq: ci=i; break
                    else: 
                        for i in v_idx: 
                            if o_clr[i]!=64: ci=i; break
                
                if ci != -1: corr = o_txt[ci]
                real = [x for x in o_txt if x]
                if real and corr:
                    data.append({'category': cat, 'question':q, 'options':real[:4], 'correct_text':corr})
            return pd.DataFrame(data)

        elif filename.lower().endswith('.xlsx'):
            wb = openpyxl.load_workbook(file_bytes, data_only=False)
            sh = wb.active
            rows = list(sh.iter_rows())
            hr = -1; cols = {'u': -1, 'q': -1, 'obj': -1}
            for r_idx, row in enumerate(rows[:20]):
                vs = [normalize_text(c.value) for c in row]
                if any('سؤال' in x for x in vs) and any('وحده' in x for x in vs):
                    hr = r_idx
                    for c, v in enumerate(vs):
                        if 'وحده' in v: cols['u'] = c
                        elif 'سؤال' in v: cols['q'] = c
                        elif 'هدف' in v: cols['obj'] = c
                    break
            if hr == -1: return pd.DataFrame()
            
            s_opt = min(cols['u'], cols['q']) + 1; e_opt = max(cols['u'], cols['q'])
            opt_cols = list(range(s_opt, e_opt))
            
            data = []
            for row in rows[hr+1:]:
                try:
                    u = normalize_text(row[cols['u']].value)
                    q = str(row[cols['q']].value if row[cols['q']].value else "").strip()
                except: continue
                cat = u + (str(row[cols['obj']].value) if cols['obj']!=-1 and cols['obj']<len(row) else "")
                if not u or not q: continue
                
                o_txt = []; o_is_col = []
                for ci in opt_cols:
                    if ci < len(row):
                        c = row[ci]
                        val = str(c.value if c.value else "").strip()
                        is_col = False
                        if c.fill and c.fill.start_color:
                            if c.fill.start_color.type == 'rgb' and c.fill.start_color.rgb not in ['00000000', 'FFFFFFFF', None]: is_col=True
                            elif c.fill.start_color.type == 'theme': is_col=True
                        o_txt.append(val); o_is_col.append(is_col)
                
                corr = ""; v_idx = [i for i,x in enumerate(o_txt) if x]
                found = -1
                for i in v_idx:
                    if o_is_col[i]: found=i; break
                if found!=-1: corr = o_txt[found]
                
                real = [x for x in o_txt if x]
                if real and corr:
                    data.append({'category': cat, 'question':q, 'options':real[:4], 'correct_text':corr})
            return pd.DataFrame(data)
            
    except Exception as e:
        return pd.DataFrame() # Return empty on error
    return pd.DataFrame()

def get_master_pattern(file_bytes, limit=30):
    if not file_bytes: return [random.randint(0,3) for _ in range(limit)]
    try:
        file_bytes.seek(0)
        wb = openpyxl.load_workbook(file_bytes, data_only=False); sh = wb.active
        pat = []
        for row in sh.iter_rows():
            found = -1
            for cell in row:
                is_col = False
                if cell.fill and cell.fill.start_color:
                    if cell.fill.start_color.type == 'rgb' and cell.fill.start_color.rgb not in ['00000000', 'FFFFFFFF', None]: is_col=True
                    elif cell.fill.start_color.type == 'theme': is_col=True
                if is_col and cell.value:
                    txt = str(cell.value).strip()
                    if 'أ' in txt or 'ا' in txt: found=0
                    elif 'ب' in txt: found=1
                    elif 'ج' in txt: found=2
                    elif 'د' in txt: found=3
                    if found!=-1: pat.append(found); break
        while len(pat) < limit: pat.append(random.randint(0,3))
        return pat
    except: return [random.randint(0,3) for _ in range(limit)]

def generate_exam(df, total):
    if df.empty: return pd.DataFrame()
    grp = df.groupby('category'); cats = list(grp.groups.keys())
    base = total // len(cats); rem = total % len(cats)
    sel = []; random.shuffle(cats)
    for i, cat in enumerate(cats):
        q = base + (1 if i < rem else 0)
        g = grp.get_group(cat)
        sel.append(g.sample(n=q) if len(g)>=q else g)
    res = pd.concat(sel)
    if len(res) < total:
        rem_df = df[~df.index.isin(res.index)]
        need = total - len(res)
        if not rem_df.empty: res = pd.concat([res, rem_df.sample(n=need) if len(rem_df)>=need else rem_df])
    return res.sample(frac=1).reset_index(drop=True)

# =========================================================
#  🖥️ واجهة المستخدم (الاحترافية)
# =========================================================

st.title("🚀 نظام توليد الاختبارات (Enterprise)")
st.info("نظام مصمم للتعامل مع آلاف الملفات بدقة وسرعة. تأكد من أن ملفات الإكسل تحتوي على أعمدة 'سؤال' و 'وحدة' وملونة للإجابة.")

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### 📊 لوحة التحكم")
    num_questions = st.number_input("عدد الأسئلة في النموذج", min_value=5, max_value=100, value=30)
    
    st.markdown("### 📝 السجل الحي")
    log_box = st.empty()
    logs = []
    def log(msg):
        logs.append(f"⏱️ {time.strftime('%H:%M:%S')} - {msg}")
        log_box.markdown(f'<div class="log-box">{"<br>".join(logs[-15:])}</div>', unsafe_allow_html=True)

with col2:
    st.markdown("### 1️⃣ رفع بنوك الأسئلة (Excel)")
    uploaded_banks = st.file_uploader("اختر كل ملفات البنوك مرة واحدة", type=['xls', 'xlsx'], accept_multiple_files=True)

st.markdown("---")
st.markdown("### 2️⃣ إعدادات النماذج الستة")

models_conf = [
    {"id": "am1", "name": "ا_صباحي", "folder": "صباحي"},
    {"id": "am2", "name": "ب_صباحي", "folder": "صباحي"},
    {"id": "pm1", "name": "ا_مسائي", "folder": "مسائي"},
    {"id": "pm2", "name": "ب_مسائي", "folder": "مسائي"},
    {"id": "sec1", "name": "ا_دور_ثاني", "folder": "دور_ثاني"},
    {"id": "sec2", "name": "ب_دور_ثاني", "folder": "دور_ثاني"},
]

model_files = {}
cols = st.columns(3) # 3 كروت في الصف الواحد

for i, m in enumerate(models_conf):
    with cols[i % 3]:
        with st.expander(f"📌 {m['name']}", expanded=True):
            t = st.file_uploader("التمبلت (Word)", type=['docx'], key=f"t_{m['id']}")
            k = st.file_uploader("المفتاح (Excel)", type=['xlsx'], key=f"k_{m['id']}")
            model_files[m['name']] = {"folder": m['folder'], "t": t, "k": k}

st.markdown("---")

# --- زر التنفيذ العملاق ---
if st.button("بدء المعالجة وإنشاء الملفات ⚡", use_container_width=True):
    if not uploaded_banks:
        st.error("❌ من فضلك ارفع بنوك الأسئلة أولاً!")
    else:
        try:
            # 1. تجميع الأسئلة في الذاكرة
            log("بدء قراءة بنوك الأسئلة...")
            all_dfs = []
            valid_files = 0
            
            progress_bar = st.progress(0)
            
            for i, f in enumerate(uploaded_banks):
                df = read_question_bank(f, f.name)
                if not df.empty:
                    all_dfs.append(df)
                    valid_files += 1
                progress_bar.progress((i + 1) / len(uploaded_banks))
            
            if not all_dfs:
                st.error("❌ لم يتم العثور على أي أسئلة صالحة! تأكد من وجود أعمدة 'سؤال' و 'وحدة' وتلوين الإجابات.")
                log("فشل: الملفات فارغة أو غير متوافقة.")
            else:
                BIG_DF = pd.concat(all_dfs).reset_index(drop=True)
                log(f"✅ تم تحميل {len(BIG_DF)} سؤال من {valid_files} ملفات.")
                
                # 2. إنشاء الملفات
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    
                    total_ops = len(model_files)
                    for i, (m_name, files_data) in enumerate(model_files.items()):
                        log(f"جاري إنشاء: {m_name}...")
                        
                        # أ) المفتاح
                        pattern = get_master_pattern(files_data['k'], num_questions)
                        
                        # ب) التوليد
                        exam = generate_exam(BIG_DF, num_questions)
                        if len(exam) > num_questions: exam = exam.iloc[:num_questions]
                        exam = exam.reset_index(drop=True)
                        
                        # ج) الوورد
                        doc = Document(files_data['t']) if files_data['t'] else Document()
                        set_section_rtl(doc.sections[0])
                        doc.add_paragraph("")
                        sub = doc.add_paragraph(); fix_paragraph(sub)
                        set_font(sub.add_run("اختر الإجابة الصحيحة :"), 14, True)
                        
                        for idx, row in enumerate(exam.to_dict('records')):
                            t_idx = pattern[idx % len(pattern)]
                            opts = force_align_options(row['options'], row['correct_text'], t_idx)
                            add_question(doc, idx+1, row['question'], opts)
                        
                        # د) الحفظ في الذاكرة والضغط
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zf.writestr(f"{files_data['folder']}/{m_name}.docx", doc_io.getvalue())
                        
                        progress_bar.progress((i + 1) / total_ops)
                
                # 3. التنظيف والتحميل
                del BIG_DF
                del all_dfs
                gc.collect() # تنظيف الرامات
                
                log("تم الانتهاء! الملف جاهز للتحميل.")
                st.markdown('<div class="success-box">✅ تم إنشاء جميع النماذج بنجاح! اضغط للتحميل</div>', unsafe_allow_html=True)
                
                st.download_button(
                    label="📥 تحميل الملفات (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Final_Exam_Pack_Pro.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"حدث خطأ غير متوقع: {e}")
            log(f"Error: {e}")
