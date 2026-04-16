import streamlit as st
import pandas as pd
import re
from io import BytesIO

# =====================================================
# Page Config
# =====================================================
st.set_page_config(
    page_title="Excel Smart Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)
# =====================================================
# 2. Authentication Logic (نظام الدخول)
# =====================================================
# يمكنك تعديل اليوزرات هنا
USERS = {
    "admin": {"password": "123", "role": "admin"},
    "mali": {"password": "456", "role": "mali"},
    "Donia": {"password": "963", "role": "Donia"},
    "boda": {"password": "789", "role": "user"}
}

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

def login_page():
    st.markdown("<div class='main-title'>🔐 نظام إدارة البيانات</div>", unsafe_allow_html=True)
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.form("login_form"):
                st.subheader("تسجيل الدخول")
                username = st.text_input("اسم المستخدم")
                password = st.text_input("كلمة المرور", type="password")
                submit = st.form_submit_button("دخول")

                if submit:
                    if username in USERS and USERS[username]["password"] == password:
                        st.session_state.logged_in = True
                        st.session_state.username = username
                        st.session_state.role = USERS[username]["role"]
                        st.rerun()
                    else:
                        st.error("البيانات غير صحيحة")

# =====================================================
# Custom CSS
# =====================================================
st.markdown("""
<style>
    .main-title { font-size: 38px; font-weight: 800; text-align: center; color: #1f2937; }
    .sub-title { text-align: center; font-size: 18px; color: #6b7280; margin-bottom: 30px; }
    .card { background: #ffffff; padding: 20px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); transition: 0.3s; }
    .card:hover { transform: translateY(-5px); }
    .footer { text-align: center; color: #9ca3af; margin-top: 40px; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# =====================================================
# Arabic normalization
# =====================================================
def normalize_arabic(text):
    if not text:
        return ""
    text = str(text)
    text = re.sub(r'[^\u0600-\u06FF]', '', text)
    return text

# =====================================================
# Core Processing Logic
# =====================================================
def process_excel(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        df.columns = df.columns.str.strip()

        # إضافة أعمدة أساسية لو مش موجودة
        for col in ['Alico Name', 'GP Code', 'Parent Code']:
            if col not in df.columns:
                df[col] = ""

        # تنظيف بعض الأعمدة
        for col in ['Sup Data Source Type', 'Alico Name']:
            if col in df.columns:
                df[col] = df[col].str.replace("'", "", regex=False).fillna("").str.strip()

        # لو فيه عمود "Id" ننسخه لـ "ID Number"
        if 'Id' in df.columns and 'ID Number' not in df.columns:
            df['ID Number'] = df['Id']

    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")
        return None

    allowed_agents = ['250218', '250712', '250610', '250602', '250205', '250907']

    # توحيد التواريخ
    for col in ['Delivery Date', 'Call Date', 'Birth Day']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: pd.to_datetime(x, errors='coerce').strftime('%d/%m/%Y') 
        if pd.notna(x) and str(x).strip() != "" and pd.to_datetime(x, errors='coerce') is not pd.NaT
        else ""
)


    # قواعد الـ Confirmation Agent
    agents_rules = [
        (r'(تم\s*عمل\s*)?(كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?(ميس|مس)?\s*نشو[هةويى]\s*(بدر)?', '060'),
        (r'امضاء\s+سار[هة](\s+احمد)?', '017'),
        (r'مضاء\s+سار[هة](\s+احمد)?', '017'),
        (r'أمضاء\s+سار[هة](\s+احمد)?', '017'),
        (r'أمضاء\s+ند[ىي]', '004'),
        (r'امضاء\s+ند[ىي]', '004'),
        (r'امضاء\s+خلود', '004'),
        (r'NADA|nada', '004'),
        (r'امضاء\s+ند[يى]\s+مصطف[يى]', '004'),
    ]

    advanced_cleaning = [
        r'y\s*e\s*s\s*b\s*a\s*n\s*k',
        r'n\s*o\s*b\s*a\s*n\s*k',
        r'b\s*a\s*n\s*k',
        r'y\s*e\s*s',
        r'n\s*o'
    ]

    pricing_logic = {
        'XCV13': {'new_code': 'YCC13', 'price': '600'},
        'XVC13': {'new_code': 'YCC13', 'price': '600'},
        'XEC13': {'new_code': 'YCE13', 'price': '500'},
        'XVM25': {'new_code': 'YCC25', 'price': '1000'},
        'XVC37': {'new_code': 'YCC37', 'price': '1350'},
        'EVO13': {'new_code': 'YCE13', 'price': '500'}
    }

    sensitive_columns = ['Agent', 'Code', 'Mobile']

    for index, row in df.iterrows():
        # حفظ البيانات الحساسة
        sensitive_values = {col: str(row[col]).strip() if col in df.columns else '' for col in sensitive_columns}
        # =====================================================
        # تنظيف أعمدة التليفونات من أي مسافات قبل أو بعد الرقم
        # =====================================================
        phone_columns = ['Mobile', 'Home Phone', 'Office Phone1', 'Office Phone2', 'Fax Number']

        for col in phone_columns:
            if col in df.columns:
                value = row.get(col)
                if pd.notna(value):
                    clean_value = str(value).strip()
                    df.at[index, col] = clean_value

        comment = '' if pd.isna(row.get('Delivery Comments')) else str(row.get('Delivery Comments')).strip()

        # إزالة علامة ' من الكومنت
        comment = comment.replace("'", "")

        # ===== Sales Signature ONLY (Safe Add-on) =====
        sales_signature_pattern = r'\b(امضائي|أمضائي|امضائى|أمضائى|بأمضتى|امضاء السيلز|أمضاء السيلز)\b'

        if re.search(sales_signature_pattern, comment):
            df.at[index, 'Confirmation Agent'] = '004'
            comment = re.sub(sales_signature_pattern, '', comment).strip()

        agent_id = sensitive_values.get('Agent', '')
        current_prod = str(row.get('Product', '')).strip()
        sup_source_val = str(row.get('Sup Data Source Type', '')).strip()

        # تعيين Alico Name + تعديل GP Code وParent Code لو Rec
        if sup_source_val != "":
            df.at[index, 'Alico Name'] = sup_source_val
        else:
            apply_keywords = r'(apply|ابلاى|ابلاي|rec\s*apply|rec\s*app)'
            if re.search(apply_keywords, comment, re.IGNORECASE):
                df.at[index, 'Alico Name'] = 'Apply'
                df.at[index, 'GP Code'] = '355'
                df.at[index, 'Parent Code'] = '006'
                comment = re.sub(apply_keywords, '', comment, flags=re.IGNORECASE).strip()
            else:
                df.at[index, 'Alico Name'] = 'Rec'
                # ✨ إضافة جديدة
                df.at[index, 'GP Code'] = '250'
                df.at[index, 'Parent Code'] = '004'

        # تعديل agent في بعض الحالات
        if current_prod == 'SUP6':
            if agent_id.startswith('250'):
                agent_id = '252' + agent_id[3:]
            elif agent_id.startswith('201'):
                agent_id = '202' + agent_id[3:]
            sensitive_values['Agent'] = agent_id

        if current_prod == 'SUPP6':
            if agent_id.startswith('250'):
                agent_id = '252' + agent_id[3:]
            elif agent_id.startswith('201'):
                agent_id = '202' + agent_id[3:]
            sensitive_values['Agent'] = agent_id

        if current_prod == 'EVO13' and agent_id not in allowed_agents:
            df.at[index, 'Product'] = 'XEC13'
            current_prod = 'XEC13'
        elif current_prod == 'XVC13'and agent_id not in allowed_agents:
            df.at[index, 'Product'] = 'XCV13'
            current_prod = 'XCV13'

        # معالجة Confirmation Agent للأسماء العربية
        normalized = normalize_arabic(comment)

        if any(name in normalized for name in ['مروهمحمد', 'مروةمحمد', 'موره محمد'.replace(" ", "")]):
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '250602' else '034'
            comment = re.sub(
        r'(تم\s*عمل\s*كونفيرم\s*)?'
        r'(ا.?م.?ض.?ا.?ء?\s*)?'
        r'(مرو[هة]|موره)\s*محمد', '',
            comment, flags=re.IGNORECASE)
        elif 'مروه' in normalized or 'مروة' in normalized:
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '250602' else '034'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*مرو[هة]?', '', comment, flags=re.IGNORECASE)
        elif 'يوسف' in normalized:
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '250920' else '025'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*يوسف(\s*ماجد)?', '', comment, flags=re.IGNORECASE)
        elif 'مريهان' in normalized:
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '201120' else '050'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*مريهان', '', comment, flags=re.IGNORECASE)
        elif any(name in normalized for name in ['ماريهان', 'مريهان', 'مايهان']):
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '201120' else '050'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*(ماريهان|مريهان|مايهان)', '',comment, flags=re.IGNORECASE)
        elif any(word in normalized for word in ['فاطمهمحمود','فاطمهمسعداوي','فاطمهسعداوي']):
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '250610' else '018'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*فاطم[هة]\s*(محمود|سعداو[يى])', '', comment, flags=re.IGNORECASE)
        elif 'مروهمصطفى' in normalized or 'مروةمصطفى' in normalized:
            df.at[index, 'Confirmation Agent'] = '004' if agent_id == '201171' else '016'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*مرو[هة]\s*مصطف[يى]', '', comment, flags=re.IGNORECASE)
        elif re.search(r'سار[هةا]*\s*احمد', normalized):
            df.at[index, 'Confirmation Agent'] = '017'
            comment = re.sub(r'ا.?م.?ض.?ا.?ء?\s*سار[هةا]*\s*احمد', '', comment,flags=re.IGNORECASE)


        # تطبيق بقية قواعد الـ agents
        for pattern, code in agents_rules:
            if re.search(pattern, comment, re.IGNORECASE):
                df.at[index, 'Confirmation Agent'] = code
                comment = re.sub(pattern, '', comment, flags=re.IGNORECASE)

        # ===== التحقق من كارت سنة ونصف / 18 شهر =====
        comment_lower = comment.lower()
        if ('كارت سنة ونصف' in comment_lower 
            or 'كارت سنه ونصف' in comment_lower
            or 'كارت سنه ونص' in comment_lower
            or '18 شهر' in comment_lower
            or 'سنة ونص' in comment_lower
            or 'سنه ونص' in comment_lower
            or 'سنة ونصف' in comment_lower 
            or 'سنه ونصف' in comment_lower   
            or 'كارت 18 شهر' in comment_lower):

            sup_source = str(row.get('Sup Data Source Type', '')).lower()

            match = re.search(r'(\d{1,2})', sup_source)
            if match:
                year_num = int(match.group(1))
                if year_num <= 24:
                    if 'Product' in df.columns:
                        df.at[index, 'Product'] = 'REF18'

        # Advanced Cleaning
        for p in advanced_cleaning:
            comment = re.sub(p, '', comment, flags=re.IGNORECASE)

        df.at[index, 'Delivery Comments'] = re.sub(r'\s+', ' ', comment).strip()

        # Pricing Logic
        if agent_id in allowed_agents and current_prod in pricing_logic:
            new_info = pricing_logic[current_prod]
            df.at[index, 'Product'] = new_info['new_code']
            if 'Total Amount' in df.columns:
                df.at[index, 'Total Amount'] = new_info['price']

        # إعادة حفظ البيانات الحساسة
        for col in sensitive_columns:
            if col in df.columns:
                df.at[index, col] = sensitive_values[col]

        # التأكد من نقل ID Number
        if 'Id' in df.columns:
            df.at[index, 'ID Number'] = df.at[index, 'Id']
        # =====================================================
        # ملء GP Code و Parent Code لو فاضي – آخر خطوة
        # =====================================================
        for col, default_val in [('GP Code', '250'), ('Parent Code', '004')]:
            if col in df.columns:
                df[col] = df[col].fillna('').replace('', default_val)

    return df

# =====================================================
# 6. Main UI Control
# =====================================================
if not st.session_state.logged_in:
    login_page()
else:
    with st.sidebar:
        st.success(f"مرحباً: {st.session_state.username}")
        if st.button("تسجيل الخروج"):
            st.session_state.logged_in = False
            st.rerun()

    st.markdown("<div class='main-title'>📊 Excel Smart Processor</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-title'>Professional Dashboard – No Logic Compromised</div>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

    if uploaded_file:
        with st.spinner("⏳ Processing data..."):
            result_df = process_excel(uploaded_file)

        if result_df is not None:
            st.success("✅ Processing completed successfully")
            st.dataframe(result_df.head(15), use_container_width=True)

            def convert_to_excel(df):
                out = BytesIO()
                with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    red_format = workbook.add_format({'bg_color': '#FFC7CE'})
                    if 'Mobile' in df.columns:
                        col_idx = df.columns.get_loc('Mobile')
                        for row_num, value in enumerate(df['Mobile'], start=1):
                            if pd.notna(value) and len(str(value).strip()) > 11:
                                worksheet.write(row_num, col_idx, str(value).strip(), red_format)
                return out.getvalue()


        # تعريف الأعمدة المطلوبة لكل نوع
        ren_columns = [
            "Card Holder Name","Address","Mobile","Home Phone",
            "Office Phone1","Office Phone2","Fax Number","E-Mail","Birth Day",
            "Delivery Date","Delivery Time","Agent",
            "Delivery Comments","Call Date","District",
            "Gender","Product","Bonus Months",
            "Card Number","Confirmation Agent",
            "Alico Name","ID Number"
        ]

        new_columns = [
            "Card Holder Name","Address","Mobile","Home Phone",
            "Office Phone1","Office Phone2","Fax Number","E-Mail",
            "Birth Day","Delivery Date","Delivery Time","Agent",
            "Delivery Comments","GP Code","Parent Code",
            "Call Date","District","Gender","Product",
            "Bonus Months","Confirmation Agent",
            "ID Number","Alico Name"
        ]

        df_ren = result_df[
            result_df["Card Number"].notna() & 
            (result_df["Card Number"].astype(str).str.strip() != "")
        ].reindex(columns=ren_columns)

        df_new = result_df[
            result_df["Card Number"].isna() | 
            (result_df["Card Number"].astype(str).str.strip() == "")
        ].reindex(columns=new_columns)

        if 'new_downloaded' not in st.session_state:
            st.session_state.new_downloaded = False
        if 'ren_downloaded' not in st.session_state:
            st.session_state.ren_downloaded = False

        st.markdown("---")
        st.subheader("📥 روابط تحميل الملفات المنفصلة")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### 🆕 New Clients")
            if st.session_state.new_downloaded:
                st.success("✅ تم تحميل شيت الـ NEW")
            else:
                st.warning("⚠️ شيت الـ NEW لم يُحمل بعد")

            st.download_button(
                label="⬇️ Download NEW Sheet",
                data=convert_to_excel(df_new),
                file_name="New_Records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=lambda: st.session_state.update({"new_downloaded": True}),
                key="btn_new"
            )

        with col2:
            st.markdown("### 🔄 Renewal Clients")
            if st.session_state.ren_downloaded:
                st.success("✅ تم تحميل شيت الـ RENEWAL")
            else:
                st.warning("⚠️ شيت الـ RENEWAL لم يُحمل بعد")

            st.download_button(
                label="⬇️ Download RENEWAL Sheet",
                data=convert_to_excel(df_ren),
                file_name="Renewal_Records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=lambda: st.session_state.update({"ren_downloaded": True}),
                key="btn_ren"
            )


st.markdown("<div class='footer'>🚀 Enterprise Excel Automation Dashboard</div>", unsafe_allow_html=True)
