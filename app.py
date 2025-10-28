import os
import math
import pandas as pd
import streamlit as st
import psycopg2
from openpyxl import load_workbook
from dotenv import load_dotenv
from google.cloud import vision
import re
import base64
import cv2
import numpy as np
from PIL import Image
import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
import tempfile
import arabic_reshaper
from bidi.algorithm import get_display




# إضافات لازمة للتبويب الذكي
from rapidfuzz import process, fuzz
import time
import openpyxl

# ---- الإعدادات العامة / البيئة ----
load_dotenv()

USERNAME = "admin"
PASSWORD = "Moraqip@123"

st.set_page_config(page_title="المراقب الذكي", layout="wide")

# ---- إعداد Google Vision من secrets ----
def setup_google_vision():
    try:
        key_b64 = st.secrets["GOOGLE_VISION_KEY_B64"]
        key_bytes = base64.b64decode(key_b64)
        with open("google_vision.json", "wb") as f:
            f.write(key_bytes)
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "google_vision.json"
        return vision.ImageAnnotatorClient()
    except Exception as e:
        st.error(f"❌ لم يتم تحميل مفتاح Google Vision بشكل صحيح: {e}")
        return None

# ---- اتصال قاعدة البيانات ----
def get_conn():
    return psycopg2.connect(
        dbname=os.environ.get("DB_NAME"),
        user=os.environ.get("DB_USER"),
        password=os.environ.get("DB_PASSWORD"),
        host=os.environ.get("DB_HOST"),
        port=os.environ.get("DB_PORT"),
        sslmode=os.environ.get("DB_SSLMODE", "require")
    )

# ---- دالة تحويل الجنس ----
def map_gender(x):
    try:
        val = int(float(x))
        return "F" if val == 1 else "M"
    except:
        return "M"

# ---- تسجيل الدخول ----
def login():
    st.markdown(
        """
        <style>
        .login-container {
            display: flex;
            justify-content: center;
            align-items: flex-start; /* يرفع الصندوق لفوق */
            height: 100vh;
            padding-top: 10vh;       /* مسافة من فوق */
        }
        .login-box {
            background: #ffffff;
            padding: 1.5rem 2rem;
            border-radius: 12px;
            box-shadow: 0px 2px 12px rgba(0,0,0,0.1);
            text-align: center;
            width: 300px;
        }
        .stTextInput>div>div>input {
            text-align: center;
            font-size: 14px;
            height: 35px;
        }
        .stButton button {
            background: linear-gradient(90deg, #4e73df, #1cc88a);
            color: white;
            border-radius: 6px;
            padding: 0.4rem 0.8rem;
            font-size: 14px;
            font-weight: bold;
            transition: 0.2s;
            width: 100%;
        }
        .stButton button:hover {
            background: linear-gradient(90deg, #1cc88a, #4e73df);
            transform: scale(1.02);
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown('<div class="login-container"><div class="login-box">', unsafe_allow_html=True)

    st.markdown("### 🔑 تسجيل الدخول")
    u = st.text_input("👤 اسم المستخدم", key="login_user")
    p = st.text_input("🔒 كلمة المرور", type="password", key="login_pass")

    # ✅ كبسة واحدة تكفي
    login_btn = st.button("🚀 دخول", key="login_btn")
    if login_btn:
        if u == USERNAME and p == PASSWORD:
            st.session_state.logged_in = True
            st.rerun()   # إعادة تحميل الصفحة مباشرة
        else:
            st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")

    st.markdown('</div></div>', unsafe_allow_html=True)


# ---- تحقق من حالة الجلسة ----
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    login()
    st.stop()

# ========================== الواجهة بعد تسجيل الدخول ==========================
st.title("📊 بغداد - البحث في سجلات الناخبين")
st.markdown("سيتم البحث في قواعد البيانات باستخدام الذكاء الاصطناعي 🤖")

# ====== تبويبات ======
tab_browse, tab_single, tab_file, tab_file_name_center, tab_count, tab_check, tab_count_custom, tab_qr, tab_cards_pdf = st.tabs(
    [
        "📄 تصفّح السجلات",
        "🔍 بحث برقم",
        "📂 رفع ملف Excel",
        "🔎 البحث الذكي (اسم + مركز اقتراع)",
        "📦 عدّ البطاقات",
        "🧾 التحقق من المعلومات",
        "🧮 تحليل البيانات (COUNT)",
        "🧾 توليد QR PDF",
        "🖼️ مطابقة البطاقات → PDF"
    ]
)

# ========= تحميل خط Amiri لدعم العربية ==========
FONT_DIR = "fonts"
FONT_PATH = os.path.join(FONT_DIR, "Amiri-Regular.ttf")

# إنشاء مجلد الخطوط إذا لم يكن موجود
if not os.path.exists(FONT_DIR):
    os.makedirs(FONT_DIR)

# تسجيل الخط العربي
def register_amiri():
    try:
        pdfmetrics.registerFont(TTFont("Amiri", FONT_PATH))
        return "Amiri"
    except Exception:
        return "Helvetica"  # بديل في حال الخط ناقص

arabic_font = register_amiri()

# ========= دالة لمعالجة النص العربي ==========
def fix_arabic_text(text):
    try:
        reshaped_text = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped_text)
        return bidi_text
    except:
        return text





# ----------------------------------------------------------------------------- #
# 1) 📄 تصفّح السجلات
# ----------------------------------------------------------------------------- #
with tab_browse:
    st.subheader("📄 تصفّح السجلات مع فلاتر")

    if "page" not in st.session_state:
        st.session_state.page = 1
    if "filters" not in st.session_state:
        st.session_state.filters = {"voter": "", "name": "", "center": ""}

    colf1, colf2, colf3, colf4 = st.columns([1,1,1,1])
    with colf1:
        voter_filter = st.text_input("🔢 رقم الناخب:", value=st.session_state.filters["voter"])
    with colf2:
        name_filter = st.text_input("🧑‍💼 الاسم:", value=st.session_state.filters["name"])
    with colf3:
        center_filter = st.text_input("🏫 مركز الاقتراع:", value=st.session_state.filters["center"])
    with colf4:
        page_size = st.selectbox("عدد الصفوف", [10, 20, 50, 100], index=1)

    if st.button("🔎 تطبيق الفلاتر"):
        st.session_state.filters = {
            "voter": voter_filter.strip(),
            "name": name_filter.strip(),
            "center": center_filter.strip(),
        }
        st.session_state.page = 1

    # --- بناء شروط البحث ---
    where_clauses, params = [], []
    if st.session_state.filters["voter"]:
        where_clauses.append('CAST("رقم الناخب" AS TEXT) ILIKE %s')
        params.append(f"%{st.session_state.filters['voter']}%")
    if st.session_state.filters["name"]:
        where_clauses.append('"الاسم الثلاثي" ILIKE %s')
        params.append(f"%{st.session_state.filters['name']}%")
    if st.session_state.filters["center"]:
        where_clauses.append('"اسم مركز الاقتراع" ILIKE %s')
        params.append(f"%{st.session_state.filters['center']}%")

    where_sql = f"WHERE {' AND '.join(where_clauses)}" if where_clauses else ""

    count_sql = f'SELECT COUNT(*) FROM "Bagdad" {where_sql};'
    offset = (st.session_state.page - 1) * page_size
    data_sql = f'''
        SELECT
            "رقم الناخب","الاسم الثلاثي","الجنس","هاتف","رقم العائلة",
            "اسم مركز الاقتراع","رقم مركز الاقتراع",
            "المدينة","رقم مركز التسجيل","اسم مركز التسجيل","تاريخ الميلاد"
        FROM "Bagdad"
        {where_sql}
        ORDER BY "رقم الناخب" ASC
        LIMIT %s OFFSET %s;
    '''

    try:
        conn = get_conn()
        with conn.cursor() as cur:
            cur.execute(count_sql, params)
            total_rows = cur.fetchone()[0]

        df = pd.read_sql_query(data_sql, conn, params=params + [page_size, offset])
        conn.close()

        if not df.empty:
            df = df.rename(columns={
                "رقم الناخب": "رقم الناخب",
                "الاسم الثلاثي": "الاسم",
                "الجنس": "الجنس",
                "هاتف": "رقم الهاتف",
                "رقم العائلة": "رقم العائلة",
                "اسم مركز الاقتراع": "مركز الاقتراع",
                "رقم مركز الاقتراع": "رقم مركز الاقتراع",
                "المدينة": "المدينة",
                "رقم مركز التسجيل": "رقم مركز التسجيل",
                "اسم مركز التسجيل": "اسم مركز التسجيل",
                "تاريخ الميلاد": "تاريخ الميلاد"
            })
            df["الجنس"] = df["الجنس"].apply(map_gender)

        total_pages = max(1, math.ceil(total_rows / page_size))

        # ✅ عرض النتائج
        st.dataframe(df, use_container_width=True, height=500)

        c1, c2, c3 = st.columns([1,2,1])
        with c1:
            if st.button("⬅️ السابق", disabled=(st.session_state.page <= 1)):
                st.session_state.page -= 1
                st.experimental_rerun()
        with c2:
            st.markdown(f"<div style='text-align:center;font-weight:bold'>صفحة {st.session_state.page} من {total_pages}</div>", unsafe_allow_html=True)
        with c3:
            if st.button("التالي ➡️", disabled=(st.session_state.page >= total_pages)):
                st.session_state.page += 1
                st.experimental_rerun()

    except Exception as e:
        st.error(f"❌ خطأ أثناء التصفح: {e}")

# ----------------------------------------------------------------------------- #
# 2) 🔍 البحث برقم واحد
# ----------------------------------------------------------------------------- #
with tab_single:
    st.subheader("🔍 البحث برقم الناخب")
    voter_input = st.text_input("ادخل رقم الناخب:")
    if st.button("بحث"):
        try:
            conn = get_conn()
            query = """
                SELECT "رقم الناخب","الاسم الثلاثي","الجنس","هاتف","رقم العائلة",
                       "اسم مركز الاقتراع","رقم مركز الاقتراع",
                       "المدينة","رقم مركز التسجيل","اسم مركز التسجيل","تاريخ الميلاد"
                FROM "Bagdad" WHERE "رقم الناخب" LIKE %s
            """
            df = pd.read_sql_query(query, conn, params=(voter_input.strip(),))
            conn.close()

            if not df.empty:
                df = df.rename(columns={
                    "رقم الناخب": "رقم الناخب",
                    "الاسم الثلاثي": "الاسم",
                    "الجنس": "الجنس",
                    "هاتف": "رقم الهاتف",
                    "رقم العائلة": "رقم العائلة",
                    "اسم مركز الاقتراع": "مركز الاقتراع",
                    "رقم مركز الاقتراع": "رقم مركز الاقتراع",
                    "المدينة": "المدينة",
                    "رقم مركز التسجيل": "رقم مركز التسجيل",
                    "اسم مركز التسجيل": "اسم مركز التسجيل",
                    "تاريخ الميلاد": "تاريخ الميلاد"
                })
                df["الجنس"] = df["الجنس"].apply(map_gender)

                st.dataframe(df, use_container_width=True, height=500)
            else:
                st.warning("⚠️ لم يتم العثور على نتائج")
        except Exception as e:
            st.error(f"❌ خطأ: {e}")

# ----------------------------------------------------------------------------- #
# 3) 📂 رفع ملف Excel (معدل مع الأرقام غير الموجودة)
# ----------------------------------------------------------------------------- #
with tab_file:
    st.subheader("📂 البحث باستخدام ملف Excel")
    uploaded_file = st.file_uploader("📤 ارفع ملف (رقم الناخب)", type=["xlsx"])
    if uploaded_file and st.button("🚀 تشغيل البحث"):
        try:
            voters_df = pd.read_excel(uploaded_file, engine="openpyxl")
            voter_col = "رقم الناخب" if "رقم الناخب" in voters_df.columns else "VoterNo"
            voters_list = voters_df[voter_col].astype(str).tolist()

            conn = get_conn()
            placeholders = ",".join(["%s"] * len(voters_list))
            query = f"""
                SELECT "رقم الناخب","الاسم الثلاثي","الجنس","هاتف","رقم العائلة",
                       "اسم مركز الاقتراع","رقم مركز الاقتراع",
                       "المدينة","رقم مركز التسجيل","اسم مركز التسجيل","تاريخ الميلاد"
                FROM "Bagdad" WHERE "رقم الناخب" IN ({placeholders})
            """
            df = pd.read_sql_query(query, conn, params=voters_list)
            conn.close()

            if not df.empty:
                df = df.rename(columns={
                    "رقم الناخب": "رقم الناخب",
                    "الاسم الثلاثي": "الاسم",
                    "الجنس": "الجنس",
                    "هاتف": "رقم الهاتف",
                    "رقم العائلة": "رقم العائلة",
                    "اسم مركز الاقتراع": "مركز الاقتراع",
                    "رقم مركز الاقتراع": "رقم مركز الاقتراع",
                    "المدينة": "المدينة",
                    "رقم مركز التسجيل": "رقم مركز التسجيل",
                    "اسم مركز التسجيل": "اسم مركز التسجيل",
                    "تاريخ الميلاد": "تاريخ الميلاد"
                })
                df["الجنس"] = df["الجنس"].apply(map_gender)

                df["رقم المندوب الرئيسي"] = ""
                df["الحالة"] = 0
                df["ملاحظة"] = ""
                df["رقم المحطة"] = 1

                df = df[["رقم الناخب","الاسم","الجنس","رقم الهاتف",
                         "رقم العائلة","مركز الاقتراع","رقم مركز الاقتراع","رقم المحطة",
                         "رقم المندوب الرئيسي","الحالة","ملاحظة"]]

                # ✅ إيجاد الأرقام غير الموجودة
                found_numbers = set(df["رقم الناخب"].astype(str).tolist())
                missing_numbers = [num for num in voters_list if num not in found_numbers]

                # عرض النتائج الموجودة
                st.dataframe(df, use_container_width=True, height=500)

                # ملف النتائج الموجودة
                output_file = "نتائج_البحث.xlsx"
                df.to_excel(output_file, index=False, engine="openpyxl")
                wb = load_workbook(output_file)
                wb.active.sheet_view.rightToLeft = True
                wb.save(output_file)
                with open(output_file, "rb") as f:
                    st.download_button("⬇️ تحميل النتائج", f,
                        file_name="نتائج_البحث.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # عرض وتحميل الأرقام غير الموجودة
                if missing_numbers:
                    st.warning("⚠️ الأرقام التالية لم يتم العثور عليها في قاعدة البيانات:")
                    st.write(missing_numbers)

                    missing_df = pd.DataFrame(missing_numbers, columns=["الأرقام غير الموجودة"])
                    miss_file = "missing_numbers.xlsx"
                    missing_df.to_excel(miss_file, index=False, engine="openpyxl")
                    with open(miss_file, "rb") as f:
                        st.download_button("⬇️ تحميل الأرقام غير الموجودة", f,
                            file_name="الأرقام_غير_الموجودة.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            else:
                st.warning("⚠️ لا يوجد نتائج")
        except Exception as e:
            st.error(f"❌ خطأ: {e}")
# ----------------------------------------------------------------------------- #
# 4️⃣ التبويب الرابع: البحث الذكي بالاسم + مركز الاقتراع (Bagdad)
# ----------------------------------------------------------------------------- #
with tab_file_name_center:
    st.subheader("🔎 البحث الذكي (اسم + مركز اقتراع) ⚡")

    file_nc = st.file_uploader("📤 ارفع ملف Excel يحتوي الاسم + اسم مركز الاقتراع", type=["xlsx"])
    run_nc = st.button("🚀 بدء البحث ومشاهدة التقدم")

    def normalize_ar(text: str) -> str:
        if not text:
            return ""
        s = str(text)
        s = s.translate(str.maketrans('', '', ''.join([
            '\u0610','\u0611','\u0612','\u0613','\u0614','\u0615','\u0616','\u0617','\u0618','\u0619','\u061A',
            '\u064B','\u064C','\u064D','\u064E','\u064F','\u0650','\u0651','\u0652','\u0653','\u0654','\u0655',
            '\u0656','\u0657','\u0658','\u0659','\u065A','\u065B','\u065C','\u065D','\u065E','\u065F','\u0670'
        ])))
        s = s.replace("ـ", "").replace(" ", "").strip()
        s = (s.replace("أ","ا").replace("إ","ا").replace("آ","ا")
             .replace("ؤ","و").replace("ئ","ي").replace("ى","ي").replace("ة","ه"))
        return s.lower()

    def normalize_fast(s):
        uniq = s.fillna("").astype(str).unique()
        mapping = {u: normalize_ar(u) for u in uniq}
        return s.fillna("").astype(str).map(mapping)

    @st.cache_data(show_spinner=False)
    def load_all_Bagdad():
        # تحميل كل البيانات التي نحتاجها من جدول Bagdad
        conn = get_conn()
        try:
            df = pd.read_sql_query(
                '''
                SELECT "رقم الناخب","الاسم الثلاثي","اسم مركز الاقتراع"
                FROM "Bagdad"
                ''',
                conn
            )
        finally:
            conn.close()
        return df

    if file_nc and run_nc:
        start = time.time()
        st.info("📦 جاري تجهيز البيانات...")

        try:
            df = pd.read_excel(file_nc, engine="openpyxl")
            df.columns = df.columns.str.strip()
            if "الاسم" not in df.columns or "اسم مركز الاقتراع" not in df.columns:
                st.error("❌ الملف يجب أن يحتوي على الأعمدة: الاسم واسم مركز الاقتراع")
                st.stop()
        except Exception as e:
            st.error(f"❌ خطأ في قراءة الملف: {e}")
            st.stop()

        # تطبيع الأسماء والمراكز
        df["__norm_name"] = normalize_fast(df["الاسم"])
        df["__norm_center"] = normalize_fast(df["اسم مركز الاقتراع"])

        # جميع بيانات بغداد
        db_all = load_all_Bagdad()
        db_all["__norm_name"] = normalize_fast(db_all["الاسم الثلاثي"])
        db_all["__norm_center"] = normalize_fast(db_all["اسم مركز الاقتراع"])

        results = []
        total = len(df)
        progress = st.progress(0)
        status = st.empty()

        for i, row in df.iterrows():
            orig_name = str(row["الاسم"])
            orig_center = str(row["اسم مركز الاقتراع"])
            norm_name = row["__norm_name"]
            norm_center = row["__norm_center"]

            # القيم الافتراضية
            match_row = {
                "الاسم في الملف": orig_name,
                "النسبة تطابق الاسم": 0,
                "اسم من القاعدة": "—",
                "تطابق المدرسة": "—",
                "نسبة تطابق المدرسة": 0,
                "اسم مركز الاقتراع في القاعدة": "—",
                "رقم الناخب": "—"
            }

            # أولاً: البحث عبر الاسم في كامل الجدول
            # نجمع كل السجلات التي تطابق الاسم بنسبة فوق حد معين
            db_names = db_all["__norm_name"].tolist()
            scores = process.cdist([norm_name], db_names, scorer=fuzz.token_sort_ratio)[0]

            # نأخذ أفضل نتيجة
            best_idx = int(scores.argmax())
            best_score = scores[best_idx]

            # حد للتطابق المقبول، مثلاً 60٪ أو 70٪ حسب دقة الاسم
            MIN_NAME_MATCH = 60

            if best_score >= MIN_NAME_MATCH:
                rec = db_all.iloc[best_idx]
                match_row["النسبة تطابق الاسم"] = round(best_score, 2)
                match_row["اسم من القاعدة"] = rec["الاسم الثلاثي"]
                match_row["اسم مركز الاقتراع في القاعدة"] = rec["اسم مركز الاقتراع"]
                match_row["رقم الناخب"] = rec["رقم الناخب"]

                # الآن نتحقق من المدرسة
                # نحسب تطابق المدرسة بين norm_center و rec["__norm_center"]
                rec_norm_center = rec["__norm_center"]
                center_score = fuzz.ratio(norm_center, rec_norm_center)
                match_row["نسبة تطابق المدرسة"] = round(center_score, 2)
                if center_score >= 80:  # حد مناسب للمدرسة
                    match_row["تطابق المدرسة"] = "✅"
                else:
                    match_row["تطابق المدرسة"] = "❌"

            else:
                # الاسم غير موجود بدرجة كافية
                match_row["تطابق المدرسة"] = "—"

            results.append(match_row)

            # تحديث التقدم
            progress.progress((i+1)/total)
            status.text(f"معالجة {i+1}/{total}")

        # تحويل النتائج وإظهارها وتحميلها
        final_df = pd.DataFrame(results)
        st.dataframe(final_df, use_container_width=True, height=400)

        out_file = "نتائج_البحث_الذكي.xlsx"
        final_df.to_excel(out_file, index=False)

        with open(out_file, "rb") as f:
            st.download_button("⬇️ تحميل النتائج", f, file_name=out_file)

        st.success(f"✅ اكتمل البحث في {time.time()- start:.1f} ثانية")

# ----------------------------------------------------------------------------- #
# 5) 📦 عدّ البطاقات (أرقام 8 خانات) + بحث في القاعدة + قائمة الأرقام غير الموجودة
# ----------------------------------------------------------------------------- #
with tab_count:
    st.subheader("📦 عدّ البطاقات (أرقام 8 خانات) — بحث في القاعدة + الأرقام غير الموجودة")

    imgs_count = st.file_uploader(
        "📤 ارفع صور الصفحات (قد تحتوي أكثر من بطاقة)",
        type=["jpg","jpeg","png"],
        accept_multiple_files=True,
        key="ocr_count"
    )

    if imgs_count and st.button("🚀 عدّ البطاقات والبحث"):
        client = setup_google_vision()
        if client is None:
            st.error("❌ خطأ في إعداد Google Vision.")
        else:
            all_numbers, number_to_files, details = [], {}, []

            for img in imgs_count:
                try:
                    content = img.read()
                    image = vision.Image(content=content)
                    response = client.text_detection(image=image)
                    texts = response.text_annotations
                    full_text = texts[0].description if texts else ""

                    # استخراج أرقام مكونة من 8 خانات فقط
                    found_numbers = re.findall(r"\b\d{8}\b", full_text)
                    for n in found_numbers:
                        all_numbers.append(n)
                        number_to_files.setdefault(n, set()).add(img.name)

                    details.append({
                        "اسم الملف": img.name,
                        "عدد البطاقات (أرقام 8 خانات)": len(found_numbers),
                        "الأرقام المكتشفة (أرقام 8 خانات فقط)": ", ".join(found_numbers) if found_numbers else "لا يوجد"
                    })

                except Exception as e:
                    st.warning(f"⚠️ خطأ أثناء معالجة صورة {img.name}: {e}")

            total_cards = len(all_numbers)
            unique_numbers = sorted(list(set(all_numbers)))

            st.success("✅ تم الاستخراج الأولي للأرقام")

            # ----------------- بحث في قاعدة البيانات عن الأرقام الموجودة -----------------
            found_df = pd.DataFrame()
            missing_list = []
            if unique_numbers:
                try:
                    conn = get_conn()
                    placeholders = ",".join(["%s"] * len(unique_numbers))
                    query = f"""
                        SELECT "رقم الناخب","الاسم الثلاثي","الجنس","هاتف","رقم العائلة",
                               "اسم مركز الاقتراع","رقم مركز الاقتراع",
                               "المدينة","رقم مركز التسجيل","اسم مركز التسجيل","تاريخ الميلاد"
                        FROM "Bagdad" WHERE "رقم الناخب" IN ({placeholders})
                    """
                    found_df = pd.read_sql_query(query, conn, params=unique_numbers)
                    conn.close()

                    if not found_df.empty:
                        found_df = found_df.rename(columns={
                            "رقم الناخب": "رقم الناخب",
                            "الاسم الثلاثي": "الاسم",
                            "الجنس": "الجنس",
                            "هاتف": "رقم الهاتف",
                            "رقم العائلة": "رقم العائلة",
                            "اسم مركز الاقتراع": "مركز الاقتراع",
                            "رقم مركز الاقتراع": "رقم مركز الاقتراع",
                            "المدينة": "المدينة",
                            "رقم مركز التسجيل": "رقم مركز التسجيل",
                            "اسم مركز التسجيل": "اسم مركز التسجيل",
                            "تاريخ الميلاد": "تاريخ الميلاد"
                        })
                        found_df["الجنس"] = found_df["الجنس"].apply(map_gender)

                        # 🧩 إضافة نفس الأعمدة مثل تبويب 📂 رفع ملف Excel
                        found_df["رقم المندوب الرئيسي"] = ""
                        found_df["الحالة"] = 0
                        found_df["ملاحظة"] = ""
                        found_df["رقم المحطة"] = 1

                        # ترتيب الأعمدة
                        found_df = found_df[[
                            "رقم الناخب","الاسم","الجنس","رقم الهاتف",
                            "رقم العائلة","مركز الاقتراع","رقم مركز الاقتراع","رقم المحطة",
                            "رقم المندوب الرئيسي","الحالة","ملاحظة"
                        ]]

                    found_numbers_in_db = set(found_df["رقم الناخب"].astype(str).tolist()) if not found_df.empty else set()
                    for n in unique_numbers:
                        if n not in found_numbers_in_db:
                            files = sorted(list(number_to_files.get(n, [])))
                            missing_list.append({"رقم_الناخب": n, "المصدر(الصور)": ", ".join(files)})

                except Exception as e:
                    st.error(f"❌ خطأ أثناء البحث في قاعدة البيانات: {e}")
            else:
                st.info("ℹ️ لم يتم العثور على أي أرقام مكوّنة من 8 خانات في الصور المرفوعة.")

            # ----------------- عرض النتائج للمستخدم -----------------
            st.markdown("### 📊 ملخص الاستخراج")
            c1, c2, c3 = st.columns(3)
            c1.metric("إجمالي الأرقام (مع التكرار)", total_cards)
            c2.metric("إجمالي الأرقام الفريدة (8 خانات)", len(unique_numbers))
            c3.metric("عدد الصور المرفوعة", len(imgs_count))

            st.markdown("### 🔎 بيانات الناخبين (الموجودة في قاعدة البيانات)")
            if not found_df.empty:
                st.dataframe(found_df, use_container_width=True, height=400)
                out_found = "بيانات_الناخبين_الموجودين.xlsx"
                found_df.to_excel(out_found, index=False, engine="openpyxl")
                wb = load_workbook(out_found)
                wb.active.sheet_view.rightToLeft = True
                wb.save(out_found)
                with open(out_found, "rb") as f:
                    st.download_button("⬇️ تحميل بيانات الناخبين الموجودة", f,
                        file_name="بيانات_الناخبين_الموجودين.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("⚠️ لم يتم العثور على أي مطابقات في القاعدة.")

            st.markdown("### ❌ الأرقام غير الموجودة في القاعدة (مع اسم الصورة)")
            if missing_list:
                missing_df = pd.DataFrame(missing_list)
                st.dataframe(missing_df, use_container_width=True)
                miss_file = "missing_numbers_with_files.xlsx"
                missing_df.to_excel(miss_file, index=False, engine="openpyxl")
                with open(miss_file, "rb") as f:
                    st.download_button("⬇️ تحميل الأرقام غير الموجودة مع المصدر", f,
                        file_name="الأرقام_غير_الموجودة_مع_المصدر.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.success("✅ لا توجد أرقام مفقودة (كل الأرقام الموجودة تم إيجادها في قاعدة البيانات).")

# ----------------------------------------------------------------------------- #
# 6) 🧾 التحقق من صحة المعلومات (بواسطة باسم)
# ----------------------------------------------------------------------------- #
with tab_check:
    st.subheader("🧾 التحقق من صحة بيانات الناخبين (بواسطة باسم ⚡)")

    st.markdown("""
    **📋 التعليمات:**
    - الملف يجب أن يحتوي الأعمدة التالية:
      1. رقم الناخب  
      2. الاسم  
      3. رقم العائلة  
      4. رقم مركز الاقتراع
    """)

    uploaded_check = st.file_uploader("📤 ارفع ملف Excel للتحقق", type=["xlsx"], key="check_file")

    if uploaded_check and st.button("🚀 بدء التحقق السريع بواسطة باسم"):
        try:
            df_check = pd.read_excel(uploaded_check, engine="openpyxl")

            required_cols = ["رقم الناخب", "الاسم", "رقم العائلة", "رقم مركز الاقتراع"]
            missing_cols = [c for c in required_cols if c not in df_check.columns]

            if missing_cols:
                st.error(f"❌ الملف ناقص الأعمدة التالية: {', '.join(missing_cols)}")
            else:
                # شريط التقدم
                progress_bar = st.progress(0, text="🤖 باسم يحضّر البيانات...")
                total_steps = 4

                # تحميل أرقام الناخبين
                df_check = df_check.astype(str)
                voter_list = df_check["رقم الناخب"].str.replace(r"\.0$", "", regex=True).tolist()
                progress_bar.progress(1/total_steps, text="📥 تحميل أرقام الناخبين...")

                # جلب البيانات من القاعدة
                conn = get_conn()
                placeholders = ",".join(["%s"] * len(voter_list))
                query = f"""
                    SELECT "رقم الناخب","الاسم الثلاثي","رقم العائلة","رقم مركز الاقتراع"
                    FROM "Bagdad"
                    WHERE "رقم الناخب" IN ({placeholders})
                """
                df_db = pd.read_sql_query(query, conn, params=voter_list)
                conn.close()

                df_db = df_db.astype(str)
                df_db["رقم الناخب"] = df_db["رقم الناخب"].str.replace(r"\.0$", "", regex=True)
                progress_bar.progress(2/total_steps, text="💾 تم جلب البيانات من القاعدة...")

                # الدمج
                merged = pd.merge(df_check, df_db, on="رقم الناخب", how="left",
                                  suffixes=("_المدخل", "_القاعدة"))

                # تصحيح القيم وتحويل None
                for col in merged.columns:
                    merged[col] = merged[col].astype(str).str.replace("None", "").str.replace("nan", "").str.strip()
                    merged[col] = merged[col].str.replace(r"\.0$", "", regex=True)

                progress_bar.progress(3/total_steps, text="🧠 جاري مقارنة البيانات...")

                # التحقق
                def match(a, b): return "✅" if a == b else "❌"
                merged["تطابق الاسم"] = merged.apply(lambda r: match(r["الاسم"], r["الاسم الثلاثي"]), axis=1)
                merged["تطابق رقم العائلة"] = merged.apply(lambda r: match(r["رقم العائلة_المدخل"], r["رقم العائلة_القاعدة"]), axis=1)
                merged["تطابق المركز"] = merged.apply(lambda r: match(r["رقم مركز الاقتراع_المدخل"], r["رقم مركز الاقتراع_القاعدة"]), axis=1)

                def overall(row):
                    if row["الاسم الثلاثي"] == "":
                        return "❌ غير موجود في القاعدة"
                    elif all(row[x] == "✅" for x in ["تطابق الاسم", "تطابق رقم العائلة", "تطابق المركز"]):
                        return "✅ مطابق"
                    else:
                        return "⚠️ اختلاف"

                merged["النتيجة النهائية"] = merged.apply(overall, axis=1)

                # ملخص النتائج
                total = len(merged)
                match_count = (merged["النتيجة النهائية"] == "✅ مطابق").sum()
                diff_count = (merged["النتيجة النهائية"] == "⚠️ اختلاف").sum()
                not_found = (merged["النتيجة النهائية"] == "❌ غير موجود في القاعدة").sum()

                st.info(f"""
                ### 📊 ملخص التحقق
                - ✅ عدد السجلات المطابقة: **{match_count}**
                - ⚠️ عدد السجلات التي تحتوي اختلاف: **{diff_count}**
                - ❌ غير موجود في القاعدة: **{not_found}**
                - 📄 إجمالي السجلات: **{total}**
                """)

                # عرض النتائج
                st.dataframe(merged, use_container_width=True, height=450)

                # تحميل
                out_file = "نتائج_التحقق_السريع.xlsx"
                merged.to_excel(out_file, index=False, engine="openpyxl")
                with open(out_file, "rb") as f:
                    st.download_button("⬇️ تحميل نتائج التحقق", f,
                        file_name=out_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                progress_bar.progress(1.0, text="✅ التحقق اكتمل!")

        except Exception as e:
            st.error(f"❌ خطأ أثناء التنفيذ: {e}")


# ----------------------------------------------------------------------------- #
# 8) 🧩 تحليل مخصص (Group & Count)
# ----------------------------------------------------------------------------- #
st.markdown("---")
st.subheader("🧩 تحليل مخصص (Group & Count)")

st.markdown("""
**📋 التعليمات:**
- ارفع ملف Excel يحتوي الأعمدة المطلوبة.  
- اختر الأعمدة التي تريد تجميع النتائج بناءً عليها (مثلاً: *رقم مركز الاقتراع + رقم مركز التسجيل*).  
- اختر العمود الذي تريد حساب عدد تكراراته (COUNT).  
- باسم سيُظهر لك عدد الصفوف (الناخبين مثلاً) ضمن كل مجموعة 👇
""")

uploaded_group = st.file_uploader("📤 ارفع ملف Excel للتحليل المخصص", type=["xlsx"], key="group_file")

if uploaded_group:
    try:
        df = pd.read_excel(uploaded_group, engine="openpyxl")
        st.success(f"✅ تم تحميل الملف ({len(df)} صف)")

        st.markdown("### 🧱 الأعمدة المتوفرة:")
        st.write(list(df.columns))

        group_cols = st.multiselect("📊 اختر الأعمدة للتجميع (Group By):", options=df.columns)
        count_col = st.selectbox("🔢 اختر العمود المراد عده (COUNT):", options=df.columns)

        if group_cols and count_col and st.button("🚀 تنفيذ التحليل المخصص"):
            progress = st.progress(0, text="🤖 باسم يحلل البيانات...")
            total_steps = 3

            # الخطوة 1️⃣ - تجهيز البيانات
            progress.progress(1/total_steps, text="🧮 تجميع البيانات...")

            # الخطوة 2️⃣ - حساب عدد الصفوف حسب الأعمدة المحددة
            grouped = df.groupby(group_cols)[count_col].count().reset_index()
            grouped = grouped.rename(columns={count_col: "عدد الصفوف"})

            progress.progress(2/total_steps, text="📊 تجهيز النتائج...")

            # الخطوة 3️⃣ - عرض وتحميل النتائج
            st.dataframe(grouped, use_container_width=True, height=450)

            # زر تحميل النتائج
            out_file = "نتائج_تحليل_مخصص.xlsx"
            grouped.to_excel(out_file, index=False, engine="openpyxl")
            with open(out_file, "rb") as f:
                st.download_button("⬇️ تحميل النتائج (Excel)", f,
                    file_name="نتائج_تحليل_مخصص.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            progress.progress(1.0, text="✅ تم التحليل بنجاح بواسطة باسم!")
    except Exception as e:
        st.error(f"❌ حدث خطأ أثناء التحليل: {e}")

# ----------------------------------------------------------------------------- #
# 9) 🧾 توليد PDF QR (24 QR في الصفحة + دعم عربي كامل + محاذاة وسط)
# ----------------------------------------------------------------------------- #


with tab_qr:
    st.subheader("🧾 توليد ملف PDF يحتوي على QR Codes")

    uploaded_qr = st.file_uploader("📤 ارفع ملف Excel فيه (الاسم - مندوب رئيسي - رمز QR)", type=["xlsx"], key="qr_file")

    if uploaded_qr:
        try:
            import pandas as pd
            df_qr = pd.read_excel(uploaded_qr, engine="openpyxl")
        except Exception as e:
            st.error(f"❌ تعذّر قراءة الملف: {e}")
            st.stop()

        # ✅ التحقق من الأعمدة
        required_cols = ["الاسم", "مندوب رئيسي", "رمز QR"]
        if not all(col in df_qr.columns for col in required_cols):
            st.error("❌ يجب أن يحتوي الملف على الأعمدة: الاسم، مندوب رئيسي، رمز QR")
            st.stop()

        # ✅ زر إنشاء PDF
        if st.button("🚀 إنشاء ملف PDF"):
            try:
                pdf_name = "qr_codes.pdf"
                c = canvas.Canvas(pdf_name, pagesize=A4)

                rows, cols = 6, 4               # ✅ 24 كود في الصفحة
                qr_size = 95                    # ✅ حجم QR ممتاز
                page_w, page_h = A4
                x_margin, y_margin = 50, 70     # ✅ مسافات خارجية
                spacing_x = 25                  # ✅ تباعد بين الأكواد أفقياً
                spacing_y = 30                  # ✅ تباعد بين الأكواد عمودياً

                count = 0
                for _, row in df_qr.iterrows():
                    name = fix_arabic_text(str(row["الاسم"]))
                    rep = fix_arabic_text(str(row['مندوب رئيسي']))
                    link = str(row["رمز QR"])

                    if count % 24 == 0 and count != 0:
                        c.showPage()

                    qr_img = qrcode.make(link)
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    qr_img.save(tmp.name)

                    row_i = (count % 24) // cols
                    col_i = (count % 24) % cols

                    x = x_margin + col_i * (qr_size + spacing_x)
                    y = page_h - y_margin - (row_i + 1) * (qr_size + spacing_y)

                    c.drawImage(tmp.name, x, y, width=qr_size, height=qr_size)

                    c.setFont(arabic_font, 11)
                    c.drawCentredString(x + qr_size/2, y - 14, name)
                    c.drawCentredString(x + qr_size/2, y - 28, rep)

                    count += 1

                c.save()

                with open(pdf_name, "rb") as f:
                    st.download_button("⬇️ تحميل ملف PDF", f, file_name=pdf_name, mime="application/pdf")

                st.success("✅ تم إنشاء ملف QR PDF بنجاح!")

            except Exception as e:
                st.error(f"❌ حدث خطأ أثناء إنشاء PDF: {e}")
# -*- coding: utf-8 -*-
# ======================= Baghdad Voters • Cards Matching → PDF =======================
import io, os, re, tempfile, datetime as _dt
from math import ceil

import numpy as np
from PIL import Image
import streamlit as st
import cv2

from rapidfuzz import fuzz

# ReportLab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import black
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# ------------------------------------ UI Setup ------------------------------------
st.set_page_config(page_title="مطابقة البطاقات → PDF", layout="wide")

st.title("بغداد - البحث في سجلات الناخبين")
st.subheader("🖼️ مطابقة البطاقات وتجهيز PDF")

st.markdown("""
**الخطوات:**
1) ارفع صورة/صور فيها بطاقات.  
2) قصّ تلقائي → OCR (إن توفّر) → تصنيف (موحّدة/ناخب) → مطابقة بالأسماء.  
3) توليد PDF بنمط: **يسار: بطاقة الناخب** | **يمين: البطاقة الموحدة** + عمود أرقام.
""")

# -------------------------------- Arabic font & shaping --------------------------------
def _register_arabic_font():
    try_order = [
        "Amiri-Regular.ttf",
        "NotoNaskhArabic-Regular.ttf",
        "DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in try_order:
        try:
            if os.path.exists(p):
                pdfmetrics.registerFont(TTFont("ArabicUI", p))
                return "ArabicUI"
        except Exception:
            pass
    return "Helvetica"  # fallback (يشتغل حتى لو بدون تشكيل/عكس)

AR_FONT = _register_arabic_font()

def fix_arabic_text(txt: str) -> str:
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(txt or ""))
    except Exception:
        return txt or ""

# -------------------------------- Google Vision (optional) -----------------------------
def setup_google_vision():
    try:
        from google.cloud import vision
        return vision.ImageAnnotatorClient()
    except Exception:
        return None

def _vision_text_from_pil(pil_img, client):
    """Run OCR with auto-rotations; returns full text or empty if fails."""
    if client is None:
        return ""
    from google.cloud import vision
    img = np.array(pil_img.convert("RGB"))
    img = cv2.resize(img, None, fx=1.5, fy=1.5)
    g = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    g = cv2.bilateralFilter(g, 11, 75, 75)
    th = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                               cv2.THRESH_BINARY, 31, 10)
    best = ""
    for k in [0, 1, 2, 3]:
        rot = np.rot90(th, k)
        buf = io.BytesIO()
        Image.fromarray(rot).save(buf, format="PNG")
        try:
            resp = client.text_detection(image=vision.Image(content=buf.getvalue()))
            if resp and resp.text_annotations:
                txt = resp.text_annotations[0].description
                if len(txt) > len(best): best = txt
        except Exception:
            pass
    return best

# -------------------------------- Name utils & classify --------------------------------
def _normalize_ar_name(text: str) -> str:
    if not text: return ""
    s = re.sub(r"[ًٌٍَُِّْـ]", "", text)
    s = s.replace("أ","ا").replace("إ","ا").replace("آ","ا")
    s = s.replace("ؤ","و").replace("ئ","ي").replace("ى","ي").replace("ة","ه")
    return re.sub(r"\s+"," ", s.strip()).lower()

def _extract_probable_name(text: str) -> str:
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    for i, l in enumerate(lines):
        if "الاسم" in l and i+1 < len(lines):
            cand = re.sub(r"[^أ-ي\s]", "", lines[i+1])
            if len(cand.strip()) > 5: return cand.strip()
    arab = [re.sub(r"[^أ-ي\s]","",l).strip() for l in lines if len(l) > 5]
    return max(arab, key=len) if arab else ""

def _classify_card(text: str) -> str:
    t = _normalize_ar_name(text)
    unified = ["البطاقه الوطنيه الموحده","الهويه","وزارة الداخليه","بطاقة وطنية","البطاقه الموحده","موحده"]
    voter   = ["بطاقة الناخب","المفوضيه","الانتخابات","المركز","المحطه","رقم الناخب","بايومتريه","البايومتريه"]
    if any(k in t for k in unified): return "unified"
    if any(k in t for k in voter):   return "voter"
    if "رقم الناخب" in t: return "voter"
    return "unified"

# -------------------------------- Card detection & cropping -----------------------------
def _order_points(pts):
    rect = np.zeros((4,2), dtype="float32")
    s = pts.sum(axis=1); diff = np.diff(pts, axis=1)
    rect[0] = pts[np.argmin(s)]; rect[2] = pts[np.argmax(s)]
    rect[1] = pts[np.argmin(diff)]; rect[3] = pts[np.argmax(diff)]
    return rect

def _four_point_transform(image, pts):
    rect = _order_points(pts); (tl, tr, br, bl) = rect
    maxW = int(max(np.linalg.norm(br-bl), np.linalg.norm(tr-tl)))
    maxH = int(max(np.linalg.norm(tr-br), np.linalg.norm(tl-bl)))
    dst = np.array([[0,0],[maxW-1,0],[maxW-1,maxH-1],[0,maxH-1]], dtype="float32")
    M = cv2.getPerspectiveTransform(rect, dst)
    return cv2.warpPerspective(image, M, (maxW, maxH))

def _split_cards_bounding(pil_img):
    img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    H, W = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (5,5), 0)
    edges = cv2.Canny(gray, 50, 150)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (5,5))
    edges = cv2.dilate(edges, k, iterations=1)
    edges = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, k, iterations=2)
    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    crops = []
    for c in contours:
        area = cv2.contourArea(c)
        if area < 0.05*W*H or area > 0.95*W*H:  # سعة مرنة
            continue
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02*peri, True)
        x, y, w, h = cv2.boundingRect(approx)
        ratio = w / float(h)
        if ratio < 0.65 or ratio > 3.1:
            continue
        if len(approx) == 4:
            warped = _four_point_transform(img, approx.reshape(4,2))
            pad = max(2, int(min(warped.shape[:2]) * 0.01))
            warped = warped[pad:-pad, pad:-pad]
            crops.append(Image.fromarray(cv2.cvtColor(warped, cv2.COLOR_BGR2RGB)))
        else:
            crop = img[y:y+h, x:x+w]
            crops.append(Image.fromarray(cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)))
    return crops if crops else [pil_img]

# -------------------------------- PDF layout (exact grid) ------------------------------
PAGE_W, PAGE_H = A4
M_LEFT, M_RIGHT, M_TOP, M_BOTTOM = 36, 36, 54, 36
TITLE = "صورة ضوئية لمستمسكات المراقبين"
ROWS_PER_PAGE = 4
NUM_COL_W = 32
GAP_TITLE = 22

TABLE_X0 = M_LEFT
TABLE_X1 = PAGE_W - M_RIGHT
TABLE_W  = TABLE_X1 - TABLE_X0
GRID_TOP = PAGE_H - M_TOP - GAP_TITLE
GRID_BOT = M_BOTTOM
GRID_H   = GRID_TOP - GRID_BOT
ROW_H    = GRID_H / ROWS_PER_PAGE

PICS_TOTAL_W = TABLE_W - NUM_COL_W
COL_W = PICS_TOTAL_W / 2.0

def _ar_center(c, x, y, text, size=12):
    c.setFont(AR_FONT, size)
    c.drawCentredString(x, y, fix_arabic_text(text))

def _box_image(c, pil_img, x, y, w, h, pad=10):
    if pil_img is None: return
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
    pil_img.save(tmp.name)
    iw, ih = pil_img.size
    max_w, max_h = max(1, w - 2*pad), max(1, h - 2*pad)
    r = min(max_w/iw, max_h/ih)
    nw, nh = iw*r, ih*r
    c.drawImage(ImageReader(tmp.name),
                x + (w - nw)/2, y + (h - nh)/2,
                nw, nh, preserveAspectRatio=True, mask='auto')

def _draw_page(c, page_pairs):
    _ar_center(c, PAGE_W/2, PAGE_H - M_TOP + 8, TITLE, size=14)
    _ar_center(c, TABLE_X0 + (COL_W/2),         GRID_TOP + 12,
               "صورة من الوجه الأول لبطاقة الناخب البايومترية", size=11)
    _ar_center(c, TABLE_X0 + COL_W + (COL_W/2), GRID_TOP + 12,
               "صورة من الوجه الأول للبطاقة الموحدة", size=11)

    c.setStrokeColor(black)
    c.setLineWidth(1.2)
    c.rect(TABLE_X0,                GRID_BOT, PICS_TOTAL_W, GRID_H)       # إطار عمودي الصور
    c.rect(TABLE_X0 + PICS_TOTAL_W, GRID_BOT, NUM_COL_W,   GRID_H)        # عمود الأرقام
    c.line(TABLE_X0 + COL_W,        GRID_BOT, TABLE_X0 + COL_W, GRID_TOP) # فاصل عمودي

    c.setLineWidth(0.6)
    for i in range(1, ROWS_PER_PAGE):
        y = GRID_TOP - i*ROW_H
        c.line(TABLE_X0, y, TABLE_X0 + PICS_TOTAL_W + NUM_COL_W, y)

    for r in range(ROWS_PER_PAGE):
        y0 = GRID_TOP - (r+1)*ROW_H
        _ar_center(c, TABLE_X0 + PICS_TOTAL_W + (NUM_COL_W/2),
                   y0 + ROW_H/2 - 6, str(r+1), size=13)

        rec = page_pairs[r] if r < len(page_pairs) else {"unified": None, "voter": None}
        voter_img   = rec.get("voter", {}).get("img")
        unified_img = rec.get("unified", {}).get("img")

        pad = 8
        w_box, h_box = COL_W - 2*pad, ROW_H - 2*pad
        voter_x, voter_y = TABLE_X0 + pad,         y0 + pad   # يسار = بطاقة الناخب
        uni_x,   uni_y   = TABLE_X0 + COL_W + pad, y0 + pad   # يمين = البطاقة الموحدة

        _box_image(c, voter_img,   voter_x, voter_y, w_box, h_box)
        _box_image(c, unified_img, uni_x,   uni_y,   w_box, h_box)

def generate_cards_pdf_bytes(pairs):
    pages = ceil(len(pairs) / ROWS_PER_PAGE) if pairs else 1
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    for p in range(pages):
        chunk = pairs[p*ROWS_PER_PAGE:(p+1)*ROWS_PER_PAGE]
        _draw_page(c, chunk)
        c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

# -------------------------------- Session state & uploader -----------------------------
if "cards_files" not in st.session_state:
    st.session_state.cards_files = []   # [{'name':..., 'bytes':...}]
if "cards_pdf_bytes" not in st.session_state:
    st.session_state.cards_pdf_bytes = b""
    st.session_state.cards_pdf_name  = "matched_cards.pdf"

up = st.file_uploader("📤 ارفع الصور (JPG/PNG)", type=["jpg","jpeg","png"], accept_multiple_files=True)
if up:
    st.session_state.cards_files = [{"name": f.name, "bytes": f.read()} for f in up]
    st.success(f"تم حفظ {len(st.session_state.cards_files)} ملف/ملفات للمعالجة.")

colA, colB = st.columns(2)
with colA:
    min_match = st.slider("🔎 الحد الأدنى لنسبة تطابق الاسم (%)", 50, 100, 70, 1)
with colB:
    add_unmatched = st.checkbox("إدراج غير المطابقين في الـPDF", value=True)
show_crops = st.checkbox("👀 عرض معاينة القصّ قبل المطابقة", value=False)

st.divider()

# ------------------------------------ RUN PIPELINE ------------------------------------
if st.button("🚀 معالجة الصور وتوليد PDF", use_container_width=True):
    if not st.session_state.cards_files:
        st.warning("⚠️ رجاءً ارفع صور أولًا.")
    else:
        client = setup_google_vision()
        if client is None:
            st.info("ℹ️ Google Vision غير مهيّأ — سنُكمل بدون OCR (المطابقة قد تكون أقل دقة).")

        try:
            unified, voter = [], []
            progress = st.progress(0.0)

            for i, item in enumerate(st.session_state.cards_files):
                pil = Image.open(io.BytesIO(item["bytes"])).convert("RGB")
                crops = _split_cards_bounding(pil)

                if show_crops:
                    st.caption(f"📎 القصّات المستخرجة من: {item['name']}")
                    cols = st.columns(max(1, min(3, len(crops))))
                    for j, cp in enumerate(crops):
                        cols[j % len(cols)].image(cp, use_column_width=True)

                for j, card in enumerate(crops):
                    text = _vision_text_from_pil(card, client) if client else ""
                    name = _extract_probable_name(text) if text else ""
                    typ  = _classify_card(text) if text else "unified"

                    data = {
                        "name": name,
                        "name_norm": _normalize_ar_name(name),
                        "img": card,
                        "src": f"{item['name']}#{j+1}",
                        "text": text,
                    }
                    (unified if typ == "unified" else voter).append(data)

                progress.progress((i + 1) / len(st.session_state.cards_files))

            st.info(f"📌 تم استخراج {len(unified)} بطاقة موحدة و{len(voter)} بطاقة ناخب.")

            # -------- matching ----------
            pairs, used = [], set()
            vnames = [v["name_norm"] for v in voter]

            for u in unified:
                if not u["name_norm"]:
                    pairs.append({"unified": u, "voter": None, "score": 0}); continue
                scores = [fuzz.ratio(u["name_norm"], vn) for vn in vnames] if vnames else []
                if scores:
                    idx = int(np.argmax(scores)); score = int(scores[idx])
                else:
                    idx, score = -1, 0
                if score >= min_match and idx not in used:
                    pairs.append({"unified": u, "voter": voter[idx], "score": score}); used.add(idx)
                else:
                    pairs.append({"unified": u, "voter": None, "score": score})

            if add_unmatched:
                for i, v in enumerate(voter):
                    if i not in used:
                        pairs.append({"unified": None, "voter": v, "score": 0})

            # -------- PDF + download state ----------
            pdf_bytes = generate_cards_pdf_bytes(pairs)
            st.session_state.cards_pdf_bytes = pdf_bytes
            st.session_state.cards_pdf_name  = f"matched_cards_{_dt.datetime.now():%Y%m%d_%H%M%S}.pdf"
            st.success("✅ تم إنشاء ملف PDF النهائي.")
        except Exception as e:
            st.exception(e)

# ------------------------------------ Download Button ----------------------------------
if st.session_state.cards_pdf_bytes:
    st.download_button(
        "⬇️ تحميل PDF الناتج",
        data=st.session_state.cards_pdf_bytes,
        file_name=st.session_state.cards_pdf_name,
        mime="application/pdf",
        key="download_cards_pdf",
        use_container_width=True
    )
