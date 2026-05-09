import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import time
import pytz
from datetime import datetime

# ────────────────────────────────────────────────
#           إعداد واجهة ستريمليت
# ────────────────────────────────────────────────

st.set_page_config(
    page_title="MergeX • دمج احترافي لملفات إكسل",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# إضافة خطوط Google Fonts وتنسيقات CSS المتقدمة
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700&display=swap');

    * {
        font-family: 'Tajawal', sans-serif;
    }

    /* الخلفية العامة */
    [data-testid="stAppViewContainer"] {
        background: radial-gradient(circle at top right, #1e1b4b, #0f172a) !important;
        background-attachment: fixed !important;
    }                                              
    [data-testid="stHeader"] {
        background: transparent !important;
    }

    /* الحاوية الرئيسية (Glassmorphism) */
    .block-container {
        background: rgba(30, 41, 59, 0.7) !important;
        backdrop-filter: blur(12px) !important;
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 24px;
        padding: 3rem 2rem !important;
        margin-top: 2rem;
        box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    }

    /* العناوين */
    h1, h2, h3, .stTitle {
        background: linear-gradient(90deg, #a78bfa, #f472b6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800 !important;
        text-align: center;
        letter-spacing: -0.02em;
    }

    /* الشريط الجانبي */
    section[data-testid="stSidebar"] {
        background: rgba(15, 23, 42, 0.95) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }

    /* عناصر التحكم */
    .stSelectbox label, .stFileUploader label {
        color: #e2e8f0 !important;
        font-size: 1.1rem !important;
        font-weight: 500 !important;
        margin-bottom: 10px !important;
    }

    /* منطقة رفع الملفات */
    [data-testid="stFileUploadDropzone"] {
        background: rgba(30, 41, 59, 0.5) !important;
        border: 2px dashed rgba(167, 139, 250, 0.4) !important;
        border-radius: 20px !important;
        transition: all 0.3s ease;
    }

    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #a78bfa !important;
        background: rgba(167, 139, 250, 0.05) !important;
    }

    /* الأزرار */
    .stButton > button, .stDownloadButton > button {
        width: 100%;
        background: linear-gradient(135deg, #7c3aed 0%, #4f46e5 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.75rem 1.5rem !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06) !important;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-2px) scale(1.02);
        box-shadow: 0 20px 25px -5px rgba(124, 58, 237, 0.3) !important;
        background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%) !important;
    }

    /* حالة النجاح والتنبيهات */
    .stAlert {
        background: rgba(30, 41, 59, 0.8) !important;
        color: #f8fafc !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 15px !important;
    }

    /* الجداول */
    [data-testid="stDataFrame"] {
        background: rgba(15, 23, 42, 0.5) !important;
        border-radius: 12px;
        overflow: hidden;
    }

    /* تخصيص الـ Progress Bar */
    .stProgress > div > div > div > div {
        background-color: #a78bfa !important;
    }

    /* منع الـ Scrollbar من التشوه */
    ::-webkit-scrollbar {
        width: 8px;
    }
    ::-webkit-scrollbar-track {
        background: #0f172a;
    }
    ::-webkit-scrollbar-thumb {
        background: #334155;
        border-radius: 10px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: #475569;
    }

    p, .stText, [data-testid="stText"], [data-testid="stMarkdownContainer"] p {
        color: #e2e8f0 !important;
    }

    /* إصلاح عناوين الإحصائيات (مثل: إجمالي الملفات، إجمالي السجلات) */
    [data-testid="stMetricLabel"] p {
        color: #a78bfa !important; /* لون بنفسجي فاتح ليتماشى مع تصميمك */
        font-weight: 600 !important;
        font-size: 1.1rem !important;
    }

    /* إصلاح أرقام الإحصائيات نفسها (مثل: 2, 25, 0.45 ثانية) */
    [data-testid="stMetricValue"] {
        color: #ffffff !important; /* لون أبيض ساطع للرقم */
        font-weight: 800 !important;
    }

    /* تحسين رؤية قائمة الملفات المرفوعة */
    [data-testid="stFileUploaderFileName"] {
        color: #ffffff !important; /* اللون الأبيض لاسم الملف */
        font-weight: 500 !important;
    }

    [data-testid="stFileUploaderFileData"] {
        color: #cbd5e1 !important; /* لون رمادي فاتح جداً لتفاصيل المساحة (KB) */
    }

    /* تفتيح أيقونة الحذف (علامة X) */
    [data-testid="stFileUploaderDeleteBtn"] svg {
        fill: #ffffff !important;
    }

    /* 1. تفتيح أرقام الإحصائيات وعناوينها */
    [data-testid="stMetricLabel"] p {
        color: #a78bfa !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"] {
        color: #ffffff !important;
    }

   

    /* 3. توضيح أسماء الملفات المرفوعة */
    [data-testid="stFileUploaderFileName"] {
        color: #ffffff !important;
    }

    /* إخفاء الشريط العلوي بالكامل تماماً */
    [data-testid="stHeader"] {
        display: none !important;
    }

    /* إخفاء القائمة الجانبية تماماً "الهمبرجر منيو" */
    #MainMenu {
        display: none !important;
    }

    /* إخفاء أي تذييل للصفحة */
    footer {
        display: none !important;
    }

    /* إخفاء زرار الـ Deploy */
    .stDeployButton {
        display:none !important;
    }

    
.viewerBadge_container__1QSob {
    display: none !important;
}


#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.stDeployButton {display:none !important;}


[data-testid="stStatusWidget"] {
    visibility: hidden;
}



/* إخفاء شريط التنبيهات والحاشية الوردية تماماً */
/* --- الحل النهائي والشامل لإخفاء كل الزوائد --- */

    /* 1. إخفاء أيقونة GitHub وشريط الحالة (الجزء الذي سألت عنه) */
    [data-testid="stStatusWidget"],
    div[class*="stStatusWidget"],
    div[class^="stStatusWidget"],
    div[class*="viewerBadge"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        height: 0 !important;
        width: 0 !important;
    }

    /* 2. إخفاء الرأس (Header) بالكامل */
    header, [data-testid="stHeader"] {
        display: none !important;
    }

    /* 3. إخفاء زر Deploy */
    .stDeployButton, [data-testid="stAppDeployButton"] {
        display: none !important;
    }

    /* 4. إخفاء المنيو (القائمة الجانبية) والتذييل */
    #MainMenu, footer {
        display: none !important;
    }

    /* 5. تنظيف أي مساحات فارغة قد تظهر في الأعلى بعد الإخفاء */
    .block-container {
        padding-top: 2rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────
#               ثوابت البيانات والقوالب
# ────────────────────────────────────────────────

TEMPLATES = {
    "كارت اديب (ADIB)": {
        "file": "قالب.xlsx",
        "defaults": {
            "Agent": "201134",
            "GP Code": "201",
            "Parent Code": "134",
            "Product": "ADIB13",
            "Bonus Months": "0",
            "Confirmation Agent": "999"
        }
    },
    "كارت تكة (TAKKA)": {
        "file": "قالب.xlsx",
        "defaults": {
            "Agent": "183001",
            "GP Code": "201",
            "Parent Code": "183",
            "Product": "TKA13",
            "Confirmation Agent": "999"
        }
    }
}

FINAL_COLUMNS = [
    "Card Holder Name", "Address", "Mobile", "Home Phone", "Office Phone1", "Office Phone2",
    "Fax Number", "E-Mail", "Birth Day", "Delivery Date", "Delivery Time", "Agent",
    "Delivery Comments", "GP Code", "Parent Code", "Call Date", "District", "Gender",
    "Product", "Bonus Months", "Confirmation Agent", "ID Number", "Alico Name"
]

COLUMN_MAPPING = {
    "Customer Name": "Card Holder Name", "Customer Name ": "Card Holder Name",
    "customer name": "Card Holder Name", " Customer Name": "Card Holder Name",
    "CUSTOMER NAME": "Card Holder Name", "CustomerName": "Card Holder Name",
    "Customer name": "Card Holder Name", "Cust Name": "Card Holder Name",
    "Name": "Card Holder Name", "الاسم": "Card Holder Name", "اسم العميل": "Card Holder Name",
    "AM or PM": "Delivery Time", "AM/PM": "Delivery Time", "AM": "Delivery Time",
    "PM": "Delivery Time", "Delivery Time": "Delivery Time", "AM OR PM": "Delivery Time",
    "Home phone": "Home Phone", "HOME PHONE": "Home Phone", "HomePhone": "Home Phone",
    "Home Tel": "Home Phone", "HOME NUM": "Home Phone",
}

# ────────────────────────────────────────────────
#               وظائف المعالجة الذكية
# ────────────────────────────────────────────────

@st.cache_resource
def load_excel_template(template_name):
    """تحميل القالب مرة واحدة وتخزينه في الذاكرة لتسريع العمليات"""
    try:
        wb = load_workbook(template_name)
        return wb
    except Exception as e:
        st.error(f"❌ خطأ فادح في تحميل القالب: {e}")
        return None

def process_single_file(file, mapping, final_cols, defaults):
    """معالجة ملف إكسل واحد بكفاءة عالية"""
    processed_sheets = []
    try:
        # استخدام pd.ExcelFile كمدير سياق للتعامل الأفضل مع الذاكرة
        with pd.ExcelFile(file) as xl:
            for sheet_name in xl.sheet_names:
                # قراءة البيانات مع تحديد أنواع البيانات لتجنب التحذيرات وتحسين الأداء
                df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)
                
                # تنظيف أسماء الأعمدة
                df.columns = df.columns.str.strip()
                df.rename(columns=mapping, inplace=True)

                # 1. فلترة: DONE في Comment (تجاهل حالة الأحرف)
                if "Comment" in df.columns:
                    df = df[df['Comment'].astype(str).str.contains("DONE", case=False, na=False)]
                
                # 2. فلترة: استبعاد YES في Comment 2
                if "Comment 2" in df.columns:
                    df = df[df["Comment 2"].fillna("").astype(str).str.strip().str.upper() != "YES"]

                # التحقق من وجود بيانات بعد الفلترة
                if df.empty:
                    continue

                # اختيار الأعمدة الموجودة فقط
                existing_cols = [c for c in final_cols if c in df.columns]
                
                if existing_cols:
                    df_subset = df[existing_cols].copy()

                    # ملء الأعمدة المفقودة
                    for col in final_cols:
                        if col not in df_subset.columns:
                            df_subset[col] = ""

                    # إعادة الترتيب
                    df_subset = df_subset[final_cols]

                    # تطبيق القيم الافتراضية
                    for col, val in defaults.items():
                        df_subset[col] = val

                    processed_sheets.append(df_subset)
                    
        return processed_sheets
    except Exception as e:
        st.warning(f"⚠️ مشكلة في الملف {file.name}: {e}")
        return []

# ────────────────────────────────────────────────
#               واجهة المستخدم الرئيسية
# ────────────────────────────────────────────────

def main():
    st.markdown("<h1 style='margin-bottom: 0.5rem;'>MergeX Pro ⚡</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 1.1rem;'>النظام الذكي لدمج ومعالجة ملفات الإكسل بسرعة فائقة</p>", unsafe_allow_html=True)
    st.markdown("---")

    # تقسيم الواجهة إلى قسمين
    col1, col2 = st.columns([1, 2], gap="large")

    with col1:
        st.subheader("⚙️ الإعدادات")
        selected_template_key = st.selectbox(
            "اختر القالب المستهدف",
            list(TEMPLATES.keys()),
            help="سيتم استخدام هذا القالب لتنسيق البيانات النهائية"
        )
        
        template_info = TEMPLATES[selected_template_key]
        TEMPLATE_FILE = template_info["file"]

        egypt_timezone = pytz.timezone('Africa/Cairo')
        current_date = datetime.now(egypt_timezone).strftime("%d/%m/%Y")
        DEFAULT_VALS = template_info["defaults"].copy()

        DEFAULT_VALS["Call Date"] = current_date

        # تحميل القالب (مخزن مؤقتًا)
        wb_template = load_excel_template(TEMPLATE_FILE)
        
        if wb_template:
            st.success(f"✅ القالب محمل وجاهز: {selected_template_key}")
        
        st.info("💡 يتم دمج الصفوف التي تحتوي على 'DONE' في عمود التعليق فقط، وتجاهل الصفوف التي تم معالجتها مسبقاً (YES).")

    with col2:
        st.subheader("📂 رفع البيانات")
        uploaded_files = st.file_uploader(
            "اسحب ملفات الإكسل هنا",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="يمكنك اختيار ملف واحد أو عدة ملفات معاً"
        )

    # معالجة البيانات عند الرفع
   # معالجة البيانات عند الرفع
    # ────────────────────────────────────────────────
    #             معالجة البيانات تلقائياً
    # ────────────────────────────────────────────────

    if uploaded_files:
        # 1. إنشاء "بصمة" للملفات الحالية (الاسم + الحجم) للتأكد من أي تغيير
        current_files_fingerprint = [(f.name, f.size) for f in uploaded_files]
        
        # 2. التحقق: هل نحتاج لمعالجة البيانات؟ 
        # (إذا كانت أول مرة، أو إذا قام المستخدم بتغيير/إضافة ملفات)
        if (st.session_state.get('last_fingerprint') != current_files_fingerprint):
            
            all_dfs = []
            start_time = time.time()
            
            # إظهار شريط التقدم أثناء المعالجة التلقائية
            progress_container = st.container()
            with progress_container:
                st.markdown("### ⏳ جاري المعالجة التلقائية...")
                progress_bar = st.progress(0)
                
                for i, file in enumerate(uploaded_files):
                    file.seek(0)  # إعادة المؤشر لبداية الملف
                    sheets_data = process_single_file(file, COLUMN_MAPPING, FINAL_COLUMNS, DEFAULT_VALS)
                    all_dfs.extend(sheets_data)
                    progress_bar.progress((i + 1) / len(uploaded_files))
                
                end_time = time.time()

            # حفظ النتيجة والبصمة في الذاكرة
            if all_dfs:
                st.session_state['processed_data'] = pd.concat(all_dfs, ignore_index=True)
                st.session_state['process_time'] = end_time - start_time
                st.session_state['last_fingerprint'] = current_files_fingerprint
                st.rerun() # إعادة تشغيل سريعة لتنظيف واجهة الـ progress bar بعد الانتهاء
            else:
                st.session_state['processed_data'] = None
                st.warning("⚠️ لم يتم العثور على بيانات مطابقة للشروط.")

        # 3. عرض النتائج من الذاكرة (هنا السرعة القصوى في المعاينة)
        if st.session_state.get('processed_data') is not None:
            combined_df = st.session_state['processed_data']
            
            # عرض ملخص المعالجة
            st.markdown("---")
            m_col1, m_col2, m_col3 = st.columns(3)
            m_col1.metric("إجمالي الملفات", len(uploaded_files))
            m_col2.metric("إجمالي السجلات", len(combined_df))
            m_col3.metric("وقت التنفيذ", f"{st.session_state['process_time']:.2f} ثانية")

            # المعاينة ستفتح فوراً لأنها لا تعيد معالجة أي شيء
            with st.expander("👁️ معاينة البيانات المدمجة (أول 100 سجل)"):
                st.dataframe(combined_df.head(100), use_container_width=True)

            # زر الحفظ والتحميل النهائي
            st.markdown("### 💾 الخطوة النهائية")
            if st.button("🚀 توليد وتجهيز الملف للتحميل", key="generate_file"):
                with st.spinner("جاري إنشاء ملف الإكسل..."):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        combined_df.to_excel(writer, index=False, sheet_name="Merged_Data")
                    
                    st.download_button(
                        label="📥 اضغط هنا لتحميل الملف النهائي",
                        data=output.getvalue(),
                        file_name=f"MergeX_Output_{int(time.time())}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.balloons()

    # تذييل الصفحة
    st.markdown("""
        <div style='text-align: center; margin-top: 5rem; color: #64748b; font-size: 0.9rem;'>
            MergeX Pro v2.0 • تم التطوير بكل حب لخدمة أعمالكم
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()