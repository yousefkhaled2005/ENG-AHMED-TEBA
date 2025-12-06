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

st.set_page_config(page_title="نظام الاختبارات (التشخيص)", layout="wide")

# =========================================================
#  1. دوال القراءة (محسنة لكشف الأخطاء)
# =========================================================
def normalize_text(text):
    if pd.isna(text) or str(text).strip() == "": return ""
    text = str(text).strip()
    return text.replace('ة', 'ه').replace('ى', 'ي').replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')

def clean_for_comp(text):
    return normalize_text(text).replace(" ", "")

def read_question_bank_smart(file_obj, filename):
    file_obj.seek(0)
    data = []
    
    try:
        # --- معالجة XLSX (الحديث) ---
        if filename.lower().endswith('.xlsx'):
            wb = openpyxl.load_workbook(file_obj, data_only=False) # لقراءة الألوان
            # نسخة لقراءة القيم المحسوبة (لو فيه معادلات)
            file_obj.seek(0)
            wb_data = openpyxl.load_workbook(file_obj, data_only=True)
            
            sh = wb.active
            sh_data = wb_data.active
            
            rows = list(sh.iter_rows())
            rows_data = list(sh_data.iter_rows(values_only=True))
            
            hr = -1; cols = {'u': -1, 'q': -1, 'obj': -1, 'ans': -1}
            
            # البحث عن الهيدر
            for r_idx, row in enumerate(rows_data[:25]):
                vs = [normalize_text(c) for c in row]
                if any('سؤال' in x for x in vs):
                    hr = r_idx
                    for c_idx, v in enumerate(vs):
                        if 'وحده' in v: cols['u'] = c_idx
                        elif 'سؤال' in v: cols['q'] = c_idx
                        elif 'اجابه' in v or 'صحيح' in v: cols['ans'] = c_idx
                    break
            
            if hr == -1 or cols['q'] == -1: return pd.DataFrame()
            
            s_opt = cols['q'] + 1
            opt_cols = [s_opt, s_opt+1, s_opt+2, s_opt+3]
            
            for r_idx in range(hr+1, len(rows)):
                row_val = rows_data[r_idx]
                row_fmt = rows[r_idx] # للصفات والألوان
                
                try: q = str(row_val[cols['q']] if row_val[cols['q']] else "").strip()
                except: continue
                if not q: continue
                
                u = str(row_val[cols['u']]).strip() if cols['u']!=-1 and row_val[cols['u']] else "عام"
                
                o_txt = []; o_is_col = []
                for ci in opt_cols:
                    if ci < len(row_val):
                        val = str(row_val[ci] if row_val[ci] else "").strip()
                        # فحص اللون
                        is_colored = False
                        try:
                            cell = row_fmt[ci]
                            if cell.fill and cell.fill.start_color:
                                if cell.fill.start_color.type == 'rgb' and cell.fill.start_color.rgb not in ['00000000', 'FFFFFFFF', None]: is_colored = True
                                elif cell.fill.start_color.type == 'theme': is_colored = True
                        except: pass
                        o_txt.append(val); o_is_col.append(is_colored)
                    else: o_txt.append(""); o_is_col.append(False)
                
                # تحديد الإجابة
                corr = ""
                # 1. من اللون
                found_idx = -1
                for i in range(len(o_txt)):
                    if o_txt[i] and o_is_col[i]: found_idx = i; break
                
                # 2. من عمود الإجابة (إذا لم نجد لون)
                if found_idx == -1 and cols['ans'] != -1 and cols['ans'] < len(row_val):
                    ans_val = normalize_text(row_val[cols['ans']])
                    if ans_val in ['1', 'أ', 'ا']: found_idx = 0
                    elif ans_val in ['2', 'ب']: found_idx = 1
                    elif ans_val in ['3', 'ج']: found_idx = 2
                    elif ans_val in ['4', 'د']: found_idx = 3
                    # بحث بالنص
                    elif ans_val:
                         for i, txt in enumerate(o_txt):
                             if clean_for_comp(txt) == clean_for_comp(ans_val): found_idx = i; break

                if found_idx != -1: corr = o_txt[found_idx]
                
                real = [x for x in o_txt if x]
                if real: # نقبل السؤال حتى لو لم نجد إجابة لنعرضه في التشخيص
                    data.append({'category': u, 'question':q, 'options':real[:4], 'correct_text':corr})

        # --- معالجة XLS (القديم) ---
        elif filename.lower().endswith('.xls'):
            book = xlrd.open_workbook(file_contents=file_obj.read(), formatting_info=True)
            sh = book.sheet_by_index(0)
            # نفس منطق البحث عن الهيدر والأعمدة... (اختصاراً للكود هنا)
            # ... (يمكن نسخ كود xlrd السابق هنا إذا لزم الأمر)
            pass 
            
    except Exception as e: return pd.DataFrame()
    return pd.DataFrame(data)

def get_master_pattern(file_obj, limit=30):
    if not file_obj: return [random.randint(0,3) for _ in range(limit)]
    file_obj.seek(0)
    try:
        wb = openpyxl.load_workbook(file_obj, data_only=False); sh = wb.active
        pat = []
        for row in sh.iter_rows():
            found = -1
            for cell in row:
                is_col = False
                if cell.fill and cell.fill.start_color:
                    if cell.fill.start_color.type == 'rgb' and cell.fill.start_color.rgb not in ['00000000', 'FFFFFFFF', None]: is_col=True
                    elif cell.fill.start_color.type == 'theme': is_col=True
                val = str(cell.value).strip() if cell.value else ""
                if (is_col and val) or ('*' in val) or (val in ['أ','ب','ج','د']):
                    if 'أ' in val: found=0
                    elif 'ب' in val: found=1
                    elif 'ج' in val: found=2
                    elif 'د' in val: found=3
                    if found!=-1: pat.append(found); break
        while len(pat) < limit: pat.append(random.randint(0,3))
        return pat
    except: return [random.randint(0,3) for _ in range(limit)]

# =========================================================
#  2. دوال التوليد والوورد
# =========================================================
def force_align_options(options, correct_text, target_idx):
    final_opts = list(options)
    while len(final_opts) < 4: final_opts.append("---")
    
    current_idx = -1
    # تنظيف النصوص للمقارنة
    clean_corr = clean_for_comp(correct_text)
    
    for i, opt in enumerate(final_opts):
        if clean_for_comp(str(opt)) == clean_corr: current_idx = i; break
    
    status = "⚠️ لم يتم العثور على الإجابة"
    if current_idx != -1:
        if current_idx != target_idx:
            final_opts[target_idx], final_opts[current_idx] = final_opts[current_idx], final_opts[target_idx]
            status = f"✅ تم النقل ({current_idx} -> {target_idx})"
        else:
            status = f"✅ مطابق أصلاً ({current_idx})"
    else:
        # Fallback: لو مش لاقيين النص، نحطه إجباري
        if correct_text:
            final_opts[target_idx] = correct_text
            status = "⚠️ تصحيح إجباري (النص مفقود)"
            
    return final_opts[:4], status

# ... (دوال set_section_rtl, add_question, generate_exam كما هي في الكود السابق) ...
# اختصاراً للمساحة سأضعهم في التنفيذ

def generate_exam(df, total):
    if df.empty: return pd.DataFrame()
    grp = df.groupby('category'); cats = list(grp.groups.keys())
    base = total // len(cats) if cats else 0; rem = total % len(cats) if cats else 0
    sel = []; random.shuffle(cats)
    for i, cat in enumerate(cats):
        q = base + (1 if i < rem else 0)
        g = grp.get_group(cat)
        sel.append(g.sample(n=q) if len(g)>=q else g)
    res = pd.concat(sel) if sel else pd.DataFrame()
    if len(res) < total:
        left = df[~df.index.isin(res.index)]
        if not left.empty: res = pd.concat([res, left.sample(n=total-len(res)) if len(left)>=total-len(res) else left])
    return res.sample(frac=1).reset_index(drop=True)

# Word Helpers
def set_section_rtl(section):
    if not section._sectPr.find(qn('w:bidi')):
        section._sectPr.append(OxmlElement('w:bidi'))
    section.left_margin = Inches(0.4); section.right_margin = Inches(0.4)

def fix_paragraph(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.bidi = True
    p.paragraph_format.space_after = Pt(4)

def format_table(table):
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None: tbl_pr = table._tbl.add_tblPr()
    bidi = OxmlElement('w:bidiVisual'); tbl_pr.append(bidi)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'right'); tbl_pr.append(jc)
    tbl_borders = OxmlElement('w:tblBorders')
    for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{b}'); el.set(qn('w:val'), 'nil'); tbl_borders.append(el)
    tbl_pr.append(tbl_borders)
    tbl_cell_mar = OxmlElement('w:tblCellMar')
    for m in ['top', 'start', 'bottom', 'end']:
        w = OxmlElement(f'w:{m}'); w.set(qn('w:w'), '0'); w.set(qn('w:type'), 'dxa'); tbl_cell_mar.append(w)
    tbl_pr.append(tbl_cell_mar)

def set_font(run, size=14, bold=False):
    run.font.name = 'Times New Roman'; run.font.size = Pt(size); run.bold = bold
    rPr = run._element.get_or_add_rPr()
    fonts = OxmlElement('w:rFonts'); fonts.set(qn('w:ascii'), 'Times New Roman'); fonts.set(qn('w:hAnsi'), 'Times New Roman'); fonts.set(qn('w:eastAsia'), 'Times New Roman'); fonts.set(qn('w:cs'), 'Times New Roman'); rPr.append(fonts)
    if not rPr.find(qn('w:rtl')): rPr.append(OxmlElement('w:rtl'))

def add_question(doc, num, text, opts):
    t = doc.add_table(rows=2, cols=4); format_table(t)
    c = t.cell(0, 0).merge(t.cell(0, 3))
    p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.bidi = True
    set_font(p.add_run(f"{num}) {text}"), 14, True)
    chars = ['أ', 'ب', 'ج', 'د']
    for i, opt in enumerate(opts):
        if i < 4:
            c2 = t.cell(1, i); p2 = c2.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.paragraph_format.bidi = True
            set_font(p2.add_run(f"{chars[i] if i<len(chars) else '-'}. {opt}"), 14, False)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

# =========================================================
#  🖥️ واجهة المستخدم (التشخيص)
# =========================================================

st.title("🩺 نظام تشخيص وتوليد الاختبارات")
st.warning("⚠️ هذا الوضع يظهر لك البيانات التي يقرأها الكود لتكتشف سبب المشكلة.")

col1, col2 = st.columns([1, 2])
with col2:
    st.subheader("1️⃣ رفع الملفات")
    uploaded_banks = st.file_uploader("بنوك الأسئلة (xlsx فقط للتشخيص)", type=['xlsx'], accept_multiple_files=True)
    
    st.markdown("---")
    st.subheader("2️⃣ إعدادات نموذج واحد للتجربة")
    t_file = st.file_uploader("التمبلت (Word)", type=['docx'], key="t1")
    k_file = st.file_uploader("المفتاح (Excel)", type=['xlsx'], key="k1")

with col1:
    st.subheader("🔍 نتيجة الفحص")
    status_area = st.empty()

if st.button("🚀 تشغيل التشخيص والتوليد", use_container_width=True):
    if not uploaded_banks: st.error("ارفع بنك الأسئلة!")
    else:
        # 1. قراءة وعرض البيانات
        all_dfs = []
        for f in uploaded_banks:
            df = read_question_bank_smart(f, f.name)
            if not df.empty: all_dfs.append(df)
        
        if not all_dfs:
            status_area.error("❌ لم يتم قراءة أي أسئلة! تأكد من الأعمدة والألوان.")
        else:
            BIG_DF = pd.concat(all_dfs).reset_index(drop=True)
            
            # --- عرض جدول التشخيص ---
            st.write("### 📊 عينة من الأسئلة المقروءة:")
            st.dataframe(BIG_DF[['question', 'correct_text']].head(10))
            
            # فحص عدد الإجابات المفقودة
            missing_ans = BIG_DF[BIG_DF['correct_text'] == ""].shape[0]
            if missing_ans > 0:
                st.error(f"⚠️ تحذير: هناك {missing_ans} سؤال لم يتم العثور على إجابة صحيحة لهم! (راجع ملف الإكسل)")
            else:
                st.success("✅ كل الأسئلة لها إجابات صحيحة.")

            # 2. التوليد (لو فيه مفتاح)
            if k_file:
                pat = get_master_pattern(k_file)
                exam = generate_exam(BIG_DF, 30)
                
                doc = Document(t_file) if t_file else Document()
                set_section_rtl(doc.sections[0])
                
                diag_log = [] # سجل التبديل
                
                for idx, row in enumerate(exam.to_dict('records')):
                    t_idx = pat[idx % len(pat)]
                    opts, msg = force_align_options(row['options'], row['correct_text'], t_idx)
                    add_question(doc, idx+1, row['question'], opts)
                    diag_log.append(f"س{idx+1}: {msg}")
                
                # عرض سجل التبديل
                with st.expander("عرض تفاصيل ترتيب الإجابات"):
                    st.text("\n".join(diag_log))
                
                bio = io.BytesIO(); doc.save(bio)
                st.download_button("📥 تحميل النموذج التجريبي", bio.getvalue(), "Test_Exam.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
