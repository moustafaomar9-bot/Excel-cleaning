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
# Authentication Logic (نظام الدخول)
# =====================================================
USERS = {
    "admin": {"password": "123", "role": "admin"},
    "mali":  {"password": "456", "role": "mali"},
    "Donia": {"password": "963", "role": "Donia"},
    "boda":  {"password": "789", "role": "user"}
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
    .sub-title  { text-align: center; font-size: 18px; color: #6b7280; margin-bottom: 30px; }
    .card { background: #ffffff; padding: 20px; border-radius: 16px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.05); transition: 0.3s; }
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

def post_clean_comment(comment):
    """
    شيل أي بقايا من كلمة 'امضاء' بعد معالجة الكونفيرم:
    - امضاء + حرف/كلمة ≤3 حروف عربية (جزء ناقص من امضاء)
    - امضاء لوحدها
    """
    comment = re.sub(r'ا\.?م\.?ض\.?ا\.?ء?\s+[\u0600-\u06FF]{1,3}(?=\s|$)', '', comment, flags=re.IGNORECASE)
    comment = re.sub(r'\bا\.?م\.?ض\.?ا\.?ء?\b', '', comment, flags=re.IGNORECASE)
    comment = re.sub(r'\s+', ' ', comment).strip()
    return comment

# =====================================================
# Core Processing Logic
# =====================================================
def process_excel(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        df.columns = df.columns.str.strip()

        for col in ['Alico Name', 'GP Code', 'Parent Code']:
            if col not in df.columns:
                df[col] = ""

        for col in ['Sup Data Source Type', 'Alico Name']:
            if col in df.columns:
                df[col] = df[col].str.replace("'", "", regex=False).fillna("").str.strip()

        if 'Id' in df.columns and 'ID Number' not in df.columns:
            df['ID Number'] = df['Id']

    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")
        return None

    allowed_agents = ['250218', '250712', '250610', '250602', '250205', '250907']

    for col in ['Delivery Date', 'Call Date', 'Birth Day']:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: pd.to_datetime(x, errors='coerce').strftime('%d/%m/%Y')
                if pd.notna(x) and str(x).strip() != "" else ""
            )

    # ✅ agents_rules - مرتبة من الأكثر تحديداً للأقل
    agents_rules = [
        (r'(تم\s*عمل\s*)?(كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?(ميس|مس)?\s*نشو[هةويى]\s*(بدر)?', '060'),
        (r'(ا.?م.?ض.?ا.?ء?\s*|مضاء\s*|أمضاء\s*)سار[هةا]*(\s+احمد)?',                        '017'),
        (r'(ا.?م.?ض.?ا.?ء?\s*)ند[يىه]\s*مصطف[يى]',                                          '004'),
        (r'(أمضاء|امضاء)\s+ند[ىيه]',                                                          '004'),
        (r'(أمضاء|امضاء)\s+خلود',                                                             '004'),
        (r'NADA|nada',                                                                         '004'),
    ]

    advanced_cleaning = [
        r'y\s*e\s*s\s*b\s*a\s*n\s*k',
        r'n\s*o\s*b\s*a\s*n\s*k',
        r'b\s*a\s*n\s*k',
        r'y\s*e\s*s',
        r'n\s*o'
    ]

    # Pricing Logic العادي (باقي الأجنتات)
    pricing_logic = {
        'XCV13': {'new_code': 'YCC13', 'price': '600'},
        'XVC13': {'new_code': 'YCC13', 'price': '600'},
        'XEC13': {'new_code': 'YCE13', 'price': '500'},
        'XVM25': {'new_code': 'YCC25', 'price': '1000'},
        'XVC37': {'new_code': 'YCC37', 'price': '1350'},
        'EVO13': {'new_code': 'YCE13', 'price': '500'}
    }

    # Pricing Logic خاص بـ 250610
    pricing_logic_250610 = {
        'XCV13': {'new_code': 'APT13', 'price': '600'},
        'XVC13': {'new_code': 'APT13', 'price': '600'},
        'XEC13': {'new_code': 'APS13', 'price': '500'},
        'XVM25': {'new_code': 'APT25', 'price': '1000'},
        'XVC37': {'new_code': 'APT37', 'price': '1350'},
        'EVO13': {'new_code': 'APS13', 'price': '500'}
    }

    sensitive_columns = ['Agent', 'Code', 'Mobile']

    for index, row in df.iterrows():

        # حفظ البيانات الحساسة
        sensitive_values = {
            col: str(row[col]).strip() if col in df.columns else ''
            for col in sensitive_columns
        }

        # تنظيف أعمدة التليفون
        phone_columns = ['Mobile', 'Home Phone', 'Office Phone1', 'Office Phone2', 'Fax Number']
        for col in phone_columns:
            if col in df.columns and pd.notna(row.get(col)):
                df.at[index, col] = str(row.get(col)).strip()

        # استخراج الكومنت وتنظيفه
        comment = '' if pd.isna(row.get('Delivery Comments')) else str(row.get('Delivery Comments')).strip()
        comment = comment.replace("'", "")

        agent_id   = sensitive_values.get('Agent', '')
        current_prod = str(row.get('Product', '')).strip()
        sup_source_val = str(row.get('Sup Data Source Type', '')).strip()

        # ===== Alico Name + GP Code + Parent Code =====
        if sup_source_val != "":
            df.at[index, 'Alico Name'] = sup_source_val
        else:
            apply_keywords = r'(apply|ابلاى|ابلاي|rec\s*apply|rec\s*app)'
            if re.search(apply_keywords, comment, re.IGNORECASE):
                df.at[index, 'Alico Name']   = 'Apply'
                df.at[index, 'GP Code']      = '355'
                df.at[index, 'Parent Code']  = '006'
                comment = re.sub(apply_keywords, '', comment, flags=re.IGNORECASE).strip()
            else:
                df.at[index, 'Alico Name']  = 'Rec'
                df.at[index, 'GP Code']     = '250'
                df.at[index, 'Parent Code'] = '004'

        # ===== تعديل Agent للمنتجات SUP =====
        for prod in ['SUP6', 'SUPP6', 'SUP3']:
            if current_prod == prod:
                if agent_id.startswith('250'):
                    agent_id = '252' + agent_id[3:]
                elif agent_id.startswith('201'):
                    agent_id = '202' + agent_id[3:]
                sensitive_values['Agent'] = agent_id

        # ===== تعديل المنتج لغير الأجنتات المسموح بيهم =====
        if current_prod == 'EVO13' and agent_id not in allowed_agents:
            df.at[index, 'Product'] = 'XEC13'
            current_prod = 'XEC13'
        elif current_prod == 'XVC13' and agent_id not in allowed_agents:
            df.at[index, 'Product'] = 'XCV13'
            current_prod = 'XCV13'

        # =====================================================
        # Confirmation Agent Logic
        # ✅ بنستخدم متغير واحد ونوقف عند أول match
        # =====================================================
        confirmation_agent = None
        normalized = normalize_arabic(comment)

        def set_agent_and_clean(agent_code, pattern_to_remove):
            """يحط الكود ويشيل الباترن من الكومنت"""
            nonlocal comment, confirmation_agent
            confirmation_agent = agent_code
            comment = re.sub(pattern_to_remove, '', comment, flags=re.IGNORECASE).strip()

        # -------------------------------------------------------
        # كل الباترنات دي بتشتغل على comment الأصلي (فيه مسافات)
        # ✅ \s* بين كل كلمة عشان يتعامل مع المسافات الزيادة
        # -------------------------------------------------------

        # 1) مروه / مروة + محمد (الأكثر تحديداً أولاً)
        if confirmation_agent is None and any(
            name in normalized for name in ['مروهمحمد', 'مروةمحمد', 'مورهمحمد']
        ):
            code = '004' if agent_id == '250602' else '034'
            set_agent_and_clean(
                code,
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?(مرو[هة]|موره)\s*محمد'
            )

        # 2) مروه / مروة + مصطفى (قبل مروه لوحدها عشان أكثر تحديداً)
        if confirmation_agent is None and (
            'مروهمصطفى' in normalized or 'مروةمصطفى' in normalized
        ):
            code = '004' if agent_id == '201171' else '016'
            set_agent_and_clean(
                code,
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?مرو[هة]\s*مصطف[يى]'
            )

        # 3) مروه / مروة (بدون محمد أو مصطفى)
        if confirmation_agent is None and ('مروه' in normalized or 'مروة' in normalized):
            code = '004' if agent_id == '250602' else '034'
            set_agent_and_clean(
                code,
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?مرو[هة]'
            )

        # 4) يوسف
        if confirmation_agent is None and 'يوسف' in normalized:
            code = '004' if agent_id == '250920' else '025'
            set_agent_and_clean(
                code,
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?يوسف(\s*ماجد)?'
            )

        # 4) مريهان / ماريهان / مايهان
        if confirmation_agent is None and any(
            name in normalized for name in ['مريهان', 'ماريهان', 'مايهان']
        ):
            code = '004' if agent_id == '201120' else '050'
            set_agent_and_clean(
                code,
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?(ماريهان|مريهان|مايهان)'
            )

        # 6) فاطمه / فاطمة محمود / سعداوي
        if confirmation_agent is None and any(
            word in normalized for word in ['فاطمهمحمود', 'فاطمةمحمود', 'فاطمهمسعداوي', 'فاطمهسعداوي',
                                             'فاطمةمسعداوي', 'فاطمةسعداوي']
        ):
            code = '004' if agent_id == '250610' else '018'
            set_agent_and_clean(
                code,
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?فاطم[هة]\s*(محمود|مسعداو[يى]|سعداو[يى])'
            )

        # 7) ساره احمد
        if confirmation_agent is None and re.search(r'سار[هةا]*\s*احمد', normalized):
            set_agent_and_clean(
                '017',
                r'(تم\s*عمل\s*كونفيرم\s*)?(ا.?م.?ض.?ا.?ء?\s*)?سار[هةا]*\s*احمد'
            )

        # 8) امضاء السيلز / امضائي (بعد الأسماء المحددة)
        if confirmation_agent is None:
            sales_signature_pattern = r'(امضائي|أمضائي|امضائى|أمضائى|بأمضتى|امضاء\s*السيلز|أمضاء\s*السيلز)'
            if re.search(sales_signature_pattern, comment, re.IGNORECASE):
                set_agent_and_clean('004', sales_signature_pattern)

        # 9) agents_rules - ✅ بتوقف عند أول match
        if confirmation_agent is None:
            for pattern, code in agents_rules:
                if re.search(pattern, comment, re.IGNORECASE):
                    set_agent_and_clean(code, pattern)
                    break  # ✅ مهم - وقف بعد أول match

        # تطبيق الـ confirmation agent لو اتحدد
        if confirmation_agent is not None:
            df.at[index, 'Confirmation Agent'] = confirmation_agent

        # ===== كارت سنة ونصف / 18 شهر =====
        comment_lower = comment.lower()
        if any(phrase in comment_lower for phrase in [
            'كارت سنة ونصف', 'كارت سنه ونصف', 'كارت سنه ونص',
            '18 شهر', 'سنة ونص', 'سنه ونص', 'سنة ونصف',
            'سنه ونصف', 'كارت 18 شهر'
        ]):
            sup_source = str(row.get('Sup Data Source Type', '')).lower()
            match = re.search(r'(\d{1,2})', sup_source)
            if match and int(match.group(1)) <= 24:
                if 'Product' in df.columns:
                    df.at[index, 'Product'] = 'REF18'

        # Advanced Cleaning
        for p in advanced_cleaning:
            comment = re.sub(p, '', comment, flags=re.IGNORECASE)

        # ✅ شيل أي بقايا من كلمة امضاء أو حروف ناقصة
        comment = post_clean_comment(comment)

        df.at[index, 'Delivery Comments'] = re.sub(r'\s+', ' ', comment).strip()

        # ===== Pricing Logic =====
        if agent_id in allowed_agents and current_prod in pricing_logic:
            new_info = (
                pricing_logic_250610[current_prod]
                if agent_id == '250610'
                else pricing_logic[current_prod]
            )
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

    # ✅ بره الـ loop - بس على الصفوف اللي فاضية فعلاً
    if 'GP Code' in df.columns:
        df['GP Code'] = df['GP Code'].apply(
            lambda x: '250' if str(x).strip() == '' else x
        )
    if 'Parent Code' in df.columns:
        df['Parent Code'] = df['Parent Code'].apply(
            lambda x: '004' if str(x).strip() == '' else x
        )

    return df


# =====================================================
# Main UI Control
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
                    workbook  = writer.book
                    worksheet = writer.sheets['Sheet1']
                    red_format = workbook.add_format({'bg_color': '#FFC7CE'})
                    if 'Mobile' in df.columns:
                        col_idx = df.columns.get_loc('Mobile')
                        for row_num, value in enumerate(df['Mobile'], start=1):
                            if pd.notna(value) and len(str(value).strip()) > 11:
                                worksheet.write(row_num, col_idx, str(value).strip(), red_format)
                return out.getvalue()

            # أعمدة كل شيت
            ren_columns = [
                "Card Holder Name", "Address", "Mobile", "Home Phone",
                "Office Phone1", "Office Phone2", "Fax Number", "E-Mail", "Birth Day",
                "Delivery Date", "Delivery Time", "Agent",
                "Delivery Comments", "Call Date", "District",
                "Gender", "Product", "Bonus Months",
                "Card Number", "Confirmation Agent",
                "Alico Name", "ID Number"
            ]

            new_columns = [
                "Card Holder Name", "Address", "Mobile", "Home Phone",
                "Office Phone1", "Office Phone2", "Fax Number", "E-Mail",
                "Birth Day", "Delivery Date", "Delivery Time", "Agent",
                "Delivery Comments", "GP Code", "Parent Code",
                "Call Date", "District", "Gender", "Product",
                "Bonus Months", "Confirmation Agent",
                "ID Number", "Alico Name"
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

st.markdown(
    "<div class='footer'>🚀 Enterprise Excel Automation Dashboard</div>",
    unsafe_allow_html=True
)
