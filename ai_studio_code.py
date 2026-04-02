import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz, process
import io

# --- 1. وظائف تنظيف البيانات ---
def clean_arabic_text(text):
    if pd.isna(text) or not isinstance(text, str):
        return ""
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)
    text = text.replace('ة', 'ه')
    text = text.replace('أ', 'ا')
    text = text.replace('إ', 'ا')
    text = text.replace('آ', 'ا')
    text = text.replace('عبد ', 'عبد')
    return text

def find_header_row(df):
    """البحث عن صف العناوين الفعلي"""
    for i in range(min(len(df), 15)):
        if df.iloc[i].notna().sum() > len(df.columns) * 0.4:
            return i
    return 0

# --- 2. إعداد واجهة المستخدم ---
st.set_page_config(page_title="مقارن الملفات الذكي", layout="wide")
st.title("📂 نظام مقارنة وتنظيف مصنفات الاكسل الذكي")

col1, col2 = st.columns(2)

def process_upload(uploaded_file, key):
    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        selected_sheet = st.selectbox(f"اختر الورقة ({key})", sheet_names, key=f"sheet_{key}")
        
        # قراءة الملف لتحديد الرأس
        raw_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
        header_idx = find_header_row(raw_df)
        
        # قراءة البيانات الفعلية
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, skiprows=header_idx)
        
        # عرض معلومات الشيت
        st.info(f"📊 معلومات: {df.shape[0]} صف | {df.shape[1]} عمود")
        
        col_to_compare = st.selectbox(f"عمود المقارنة الأساسي", df.columns, key=f"col_{key}")
        return df, col_to_compare, selected_sheet, uploaded_file.name
    return None, None, None, None

with col1:
    st.subheader("الملف الأول")
    file1 = st.file_uploader("تحميل الملف الأول", type=['xlsx'], key="f1")
    df1, col1_name, s1_name, f1_filename = process_upload(file1, "file1")

with col2:
    st.subheader("الملف الثاني")
    file2 = st.file_uploader("تحميل الملف الثاني", type=['xlsx'], key="f2")
    df2, col2_name, s2_name, f2_filename = process_upload(file2, "file2")

# --- 3. عملية المقارنة عند الضغط على الزر ---
if st.button("بدء عملية التنظيف والمقارنة الذكية"):
    if df1 is not None and df2 is not None:
        with st.spinner("جاري معالجة وتدقيق البيانات..."):
            
            # تحضير النسخ المنظفة للمقارنة
            df1_clean = df1.copy()
            df2_clean = df2.copy()
            
            # إنشاء عمود للمقارنة فقط (مخفي)
            df1_clean['_match_key'] = df1[col1_name].astype(str).apply(clean_arabic_text)
            df2_clean['_match_key'] = df2[col2_name].astype(str).apply(clean_arabic_text)

            matches_100, matches_75, matches_50 = [], [], []

            # قائمة مفاتيح البحث من الملف الثاني
            choices = df2_clean['_match_key'].tolist()

            for idx1, row1 in df1_clean.iterrows():
                val1 = row1['_match_key']
                if not val1 or val1 == "": continue
                
                # استخدام RapidFuzz للمقارنة
                # سنحصل على: (القيمة المطابقة، النسبة، الفهرس في الملف الثاني)
                result = process.extractOne(val1, choices, scorer=fuzz.token_sort_ratio)
                
                if result:
                    matched_val, score, idx2 = result
                    
                    # بناء سجل النتيجة لتفادي InvalidIndexError
                    res_row = {}
                    # إضافة بيانات الملف الأول ببادئة
                    for c in df1.columns:
                        res_row[f"ملف1_{c}"] = df1.iloc[idx1][c]
                    
                    # إضافة بيانات الملف الثاني ببادئة
                    for c in df2.columns:
                        res_row[f"ملف2_{c}"] = df2.iloc[idx2][c]
                    
                    # إضافة الأعمدة المطلوبة في نهاية البيانات
                    res_row['نسبة التشابه'] = f"{round(score, 2)}%"
                    res_row['أين كان التشابه'] = f"التطابق بين [{df1.iloc[idx1][col1_name]}] و [{df2.iloc[idx2][col2_name]}]"
                    res_row['المصدر'] = f"ملف: {f2_filename} | ورقة: {s2_name}"
                    
                    if score == 100:
                        matches_100.append(res_row)
                    elif score >= 75:
                        matches_75.append(res_row)
                    elif score >= 50:
                        matches_50.append(res_row)

            # تحويل القوائم إلى DataFrames
            res_df_100 = pd.DataFrame(matches_100)
            res_df_75 = pd.DataFrame(matches_75)
            res_df_50 = pd.DataFrame(matches_50)

            # إنشاء ملف الاكسل
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df1.to_excel(writer, sheet_name='الأصلية-ملف 1', index=False)
                df2.to_excel(writer, sheet_name='الأصلية-ملف 2', index=False)
                
                if not res_df_100.empty:
                    res_df_100.to_excel(writer, sheet_name='تشابه 100%', index=False)
                if not res_df_75.empty:
                    res_df_75.to_excel(writer, sheet_name='تشابه 75% فما فوق', index=False)
                if not res_df_50.empty:
                    res_df_50.to_excel(writer, sheet_name='تشابه 50% إلى 74%', index=False)
            
            st.success("✅ اكتملت المقارنة!")
            st.download_button(
                label="📥 تحميل ملف النتائج",
                data=output.getvalue(),
                file_name="نتائج_المقارنة_الذكية.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("يرجى التأكد من رفع الملفين واختيار أعمدة المقارنة.")