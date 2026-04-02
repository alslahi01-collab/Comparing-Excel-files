import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz, process, utils
import io

# --- 1. وظائف تنظيف البيانات ---
def clean_arabic_text(text):
    if pd.isna(text) or not isinstance(text, str):
        return ""
    # إزالة المسافات الزائدة في البداية والنهاية
    text = text.strip()
    # توحيد المسافات المتعددة لتصبح مسافة واحدة
    text = re.sub(r'\s+', ' ', text)
    # توحيد التاء المربوطة والهاء
    text = text.replace('ة', 'ه')
    # معالجة الأسماء المركبة (حذف المسافة بعد "عبد")
    text = text.replace('عبد ', 'عبد')
    # إزالة الحروف المكررة (اختياري - يمكن توسيعه)
    text = re.sub(r'(.)\1+', r'\1', text) 
    return text

def find_header_row(df):
    """البحث عن صف العناوين الفعلي (يتجاهل الصفوف الفارغة في البداية)"""
    for i in range(min(len(df), 10)):
        if df.iloc[i].notna().sum() > len(df.columns) * 0.5:
            return i
    return 0

# --- 2. واجهة المستخدم ---
st.set_page_config(page_title="مقارن ملفات الاكسل الذكي", layout="wide")
st.title("📂 نظام مقارنة وتنظيف مصنفات الاكسل الذكي")

col1, col2 = st.columns(2)

def process_upload(uploaded_file, key):
    if uploaded_file:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        st.write(f"**معلومات الملف:** {len(sheet_names)} ورقة عمل")
        
        selected_sheet = st.selectbox(f"اختر الورقة ({key})", sheet_names, key=f"sheet_{key}")
        
        # قراءة أولية لتحديد الرأس
        raw_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)
        header_idx = find_header_row(raw_df)
        
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, skiprows=header_idx)
        st.info(f"📄 الشيت: {selected_sheet} | الصفوف: {df.shape[0]} | الأعمدة: {df.shape[1]}")
        
        col_to_compare = st.selectbox(f"اختر عمود المقارنة الأساسي (الأسماء)", df.columns, key=f"col_{key}")
        return df, col_to_compare, selected_sheet
    return None, None, None

with col1:
    st.subheader("الملف الأول")
    file1 = st.file_uploader("تحميل الملف الأول", type=['xlsx'], key="f1")
    df1, col1_name, s1_name = process_upload(file1, "file1")

with col2:
    st.subheader("الملف الثاني")
    file2 = st.file_uploader("تحميل الملف الثاني", type=['xlsx'], key="f2")
    df2, col2_name, s2_name = process_upload(file2, "file2")

# --- 3. عملية المقارنة ---
if st.button("بدء عملية التنظيف والمقارنة الذكية"):
    if df1 is not None and df2 is not None:
        with st.spinner("جاري معالجة البيانات..."):
            
            # تنظيف البيانات في الأعمدة المختارة
            df1_clean = df1.copy()
            df2_clean = df2.copy()
            
            df1_clean['compare_key'] = df1[col1_name].astype(str).apply(clean_arabic_text)
            df2_clean['compare_key'] = df2[col2_name].astype(str).apply(clean_arabic_text)

            matches_100 = []
            matches_75 = []
            matches_50 = []

            # خوارزمية المقارنة
            for idx1, row1 in df1_clean.iterrows():
                val1 = row1['compare_key']
                if not val1: continue
                
                # البحث عن أفضل تطابق في الملف الثاني
                best_match = process.extractOne(val1, df2_clean['compare_key'], score_cutoff=50)
                
                if best_match:
                    match_val, score, idx2 = best_match
                    matched_row = df2.iloc[idx2].copy()
                    
                    # إضافة أعمدة المعلومات الإضافية
                    matched_row['نسبة التشابه'] = f"{score}%"
                    matched_row['مصدر التشابه'] = f"شيت: {s2_name} | الملف الثاني"
                    matched_row['القيمة المطابقة'] = df1.iloc[idx1][col1_name]
                    
                    # دمج بيانات من الصفين للمقارنة البصرية
                    combined_result = pd.concat([df1.iloc[idx1], matched_row])
                    
                    if score == 100:
                        matches_100.append(combined_result)
                    elif score >= 75:
                        matches_75.append(combined_result)
                    elif score >= 50:
                        matches_50.append(combined_result)

            # تحويل النتائج لـ DataFrames
            res_100 = pd.DataFrame(matches_100)
            res_75 = pd.DataFrame(matches_75)
            res_50 = pd.DataFrame(matches_50)

            # --- 4. إنشاء ملف الاكسل المستخرج ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df1.to_excel(writer, sheet_name='البيانات الأصلية 1', index=False)
                df2.to_excel(writer, sheet_name='البيانات الأصلية 2', index=False)
                res_100.to_excel(writer, sheet_name='تطابق 100%', index=False)
                res_75.to_excel(writer, sheet_name='تطابق 75% فأكثر', index=False)
                res_50.to_excel(writer, sheet_name='تطابق 50% إلى 74%', index=False)
            
            processed_data = output.getvalue()
            
            st.success("✅ تمت عملية المقارنة بنجاح!")
            st.download_button(
                label="📥 تحميل ملف النتائج",
                data=processed_data,
                file_name="مخرجات_المقارنة.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("يرجى تحميل الملفين واختيار الأعمدة أولاً")