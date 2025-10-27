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
tab_browse, tab_single, tab_file, tab_file_name_center, tab_count, tab_check, tab_count_custom, tab_qr = st.tabs(
    [
        "📄 تصفّح السجلات",
        "🔍 بحث برقم",
        "📂 رفع ملف Excel",
        "🔎 البحث الذكي (اسم + مركز اقتراع)",
        "📦 عدّ البطاقات",
        "🧾 التحقق من المعلومات",
        "🧮 تحليل البيانات (COUNT)",
        "🧾 توليد QR PDF"  # ← التاب الجديد هنا ✅
        "🖼️ مطابقة البطاقات → PDF"   # ✅ التاب الجديدة
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

# ----------------------------------------------------------------------------- #
# 🪪 التاب الجديدة: مطابقة البطاقات وإنتاج PDF عمودين (هوية موحدة | بطاقة ناخب)
# ----------------------------------------------------------------------------- #
with tab_cards_pdf:
    st.subheader("🖼️ مطابقة البطاقات وتجهيز PDF ...")

    st.markdown("""
    **الآلية:**
    1) ارفع صورًا متعددة؛ قد تحتوي الصورة الواحدة على أكثر من بطاقة.  
    2) سيقوم النظام بقصّ البطاقات تلقائيًا، ثم استخدام OCR لاستخراج الاسم.  
    3) يحدّد النظام نوع البطاقة (البطاقة الوطنية الموحدة/بطاقة الناخب) عبر كلمات مفتاحية.  
    4) يطابق البطاقات حسب الاسم ويُنتج PDF نهائي فيه عمودان: **هوية موحدة | بطاقة ناخب**.  
    """)

    imgs_cards = st.file_uploader(
        "📤 ارفع صور البطاقات (JPG/PNG) – يمكن رفع عدة صور، والصورة قد تحتوي أكثر من بطاقة",
        type=["jpg","jpeg","png"],
        accept_multiple_files=True,
        key="pair_cards_uploader"
    )

    # ---------- أدوات مساعدة محلية (لا تغيّر شيئًا خارج هذا التاب) ----------
    def _normalize_ar(text: str) -> str:
        if not text:
            return ""
        s = str(text)
        # إزالة التشكيل
        s = s.translate(str.maketrans('', '', ''.join([
            '\u0610','\u0611','\u0612','\u0613','\u0614','\u0615','\u0616','\u0617','\u0618','\u0619','\u061A',
            '\u064B','\u064C','\u064D','\u064E','\u064F','\u0650','\u0651','\u0652','\u0653','\u0654','\u0655',
            '\u0656','\u0657','\u0658','\u0659','\u065A','\u065B','\u065C','\u065D','\u065E','\u065F','\u0670'
        ])))
        s = s.replace("ـ", "").strip()
        s = (s.replace("أ","ا").replace("إ","ا").replace("آ","ا")
               .replace("ؤ","و").replace("ئ","ي").replace("ى","ي").replace("ة","ه"))
        # إزالة المسافات الزائدة
        s = " ".join(s.split())
        return s

    def _guess_card_type(text: str) -> str:
        """يعيد 'unified' للبطاقة الوطنية الموحدة، 'voter' لبطاقة الناخب، أو '' عند عدم التحديد."""
        t = text or ""
        t_norm = _normalize_ar(t)
        # كلمات مفتاحية
        kw_unified = ["البطاقه الوطنيه الموحده", "الهويه الوطنيه", "البطاقه الموحده", "وزارة الداخليه", "الاحوال المدنيه"]
        kw_voter   = ["بطاقه الناخب", "المفوضيه", "الانتخابات", "المفوضيه العليا المستقله للانتخابات"]
        if any(k in t_norm for k in kw_unified):
            return "unified"
        if any(k in t_norm for k in kw_voter):
            return "voter"
        # تخمين إضافي: وجود كلمة "الناخب" vs "الوطنيه"
        if "الناخب" in t_norm:
            return "voter"
        if "الوطنيه" in t_norm or "الموحده" in t_norm:
            return "unified"
        return ""

    def _extract_name_from_text(text: str) -> str:
        """يحاول استخراج الاسم من نص عربي عبر اختيار أطول سطر عربي مناسب."""
        if not text:
            return ""
        # نقسّم لأسطر ونختار الأسطر العربية فقط
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        ar_lines = []
        for ln in lines:
            # سطر عربي (يحتوي على حروف المدى العربي)
            if re.search(r"[\u0600-\u06FF]{2,}", ln):
                # استبعاد الأسطر الصريحة التي تحوي كلمات نوع البطاقة
                ln_norm = _normalize_ar(ln)
                if "بطاقه" in ln_norm or "الناخب" in ln_norm or "الوطنيه" in ln_norm or "الموحده" in ln_norm:
                    continue
                # استبعاد الأسطر القصيرة جداً
                if len(ln_norm) < 6:
                    continue
                ar_lines.append(ln)
        if not ar_lines:
            return ""
        # اختر السطر الذي يحوي أكثر من كلمتين وعادةً يكون أطول/أغنى
        ar_lines.sort(key=lambda s: (len(s), s.count(" ")), reverse=True)
        return ar_lines[0]

    def _read_bytes(file) -> bytes:
        pos = file.tell()
        content = file.read()
        file.seek(pos)
        return content

    def _detect_and_crop_cards(img_bgr: np.ndarray) -> list:
        """
        يكشف مستطيلات تشبه البطاقات ويعيد قائمة صور مقصوصة (BGR).
        نهج بسيط: تحويل رمادي -> بلور -> Canny -> كونتورز -> ترشيح نسبة الطول/العرض والمساحة -> warp.
        """
        out = []
        h, w = img_bgr.shape[:2]
        gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (5,5), 0)
        edges = cv2.Canny(gray, 60, 180)
        edges = cv2.dilate(edges, np.ones((3,3), np.uint8), iterations=1)
        cnts, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        def four_point_transform(image, pts):
            rect = np.array(pts, dtype="float32")
            # ترتيب النقاط تقريبيًا (left-top, right-top, right-bottom, left-bottom)
            s = rect.sum(axis=1)
            diff = np.diff(rect, axis=1)
            tl = rect[np.argmin(s)]
            br = rect[np.argmax(s)]
            tr = rect[np.argmin(diff)]
            bl = rect[np.argmax(diff)]
            rect = np.array([tl, tr, br, bl], dtype="float32")
            (tl, tr, br, bl) = rect
            widthA = np.linalg.norm(br - bl)
            widthB = np.linalg.norm(tr - tl)
            maxWidth = int(max(widthA, widthB))
            heightA = np.linalg.norm(tr - br)
            heightB = np.linalg.norm(tl - bl)
            maxHeight = int(max(heightA, heightB))
            if maxWidth < 100 or maxHeight < 60:
                return None
            dst = np.array([
                [0, 0],
                [maxWidth - 1, 0],
                [maxWidth - 1, maxHeight - 1],
                [0, maxHeight - 1]], dtype="float32")
            M = cv2.getPerspectiveTransform(rect, dst)
            warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))
            return warped

        for c in cnts:
            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, 0.02 * peri, True)
            if len(approx) == 4:
                area = cv2.contourArea(approx)
                if area < (w*h) * 0.01:  # استبعاد القطع الصغيرة جدًا
                    continue
                # نسبة العرض إلى الارتفاع تقارب بطاقة (1.4 إلى 1.9 غالباً أفقية)
                x, y, cw, ch = cv2.boundingRect(approx)
                ar = cw / float(ch)
                if 1.2 <= ar <= 2.2 or 0.5 <= ar <= 0.9:  # أفقية أو عمودية
                    warped = four_point_transform(img_bgr, approx.reshape(4,2))
                    if warped is not None:
                        out.append(warped)
        # في حال لم يتم التقاط أي شيء، اعتبر الصورة كلها بطاقة واحدة
        if not out:
            out = [img_bgr]
        return out

    def _cv2_to_png_bytes(img_bgr: np.ndarray) -> bytes:
        img_rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(img_rgb)
        buf = io.BytesIO()
        pil.save(buf, format="PNG")
        return buf.getvalue()

    if imgs_cards and st.button("🚀 تنفيذ المطابقة وإنتاج PDF"):
        client = setup_google_vision()
        if client is None:
            st.error("❌ تعذّر تهيئة Google Vision.")
            st.stop()

        progress = st.progress(0)
        status = st.empty()

        # سنجمع البطاقات حسب النوع
        cards = {
            "unified": [],  # [(name_norm, name_raw, image_bgr)]
            "voter":  []
        }
        leftovers = []  # بطاقات لم يُعرف نوعها

        # 1) قصّ البطاقات + OCR
        total_files = len(imgs_cards)
        processed = 0

        for f in imgs_cards:
            try:
                content = _read_bytes(f)
                # قراءة الصورة
                file_bytes = np.asarray(bytearray(content), dtype=np.uint8)
                img_bgr = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
                if img_bgr is None:
                    st.warning(f"⚠️ تعذّر قراءة الصورة: {f.name}")
                    continue

                crops = _detect_and_crop_cards(img_bgr)

                for crop_bgr in crops:
                    # OCR
                    png_bytes = _cv2_to_png_bytes(crop_bgr)
                    gimg = vision.Image(content=png_bytes)
                    resp = client.text_detection(image=gimg)
                    texts = resp.text_annotations
                    full_text = texts[0].description if texts else ""

                    card_type = _guess_card_type(full_text)
                    name_raw  = _extract_name_from_text(full_text)
                    name_norm = _normalize_ar(name_raw)

                    rec = (name_norm, name_raw, crop_bgr)

                    if card_type in ("unified", "voter") and name_norm:
                        cards[card_type].append(rec)
                    else:
                        leftovers.append(rec)

            except Exception as e:
                st.warning(f"⚠️ خطأ أثناء معالجة {f.name}: {e}")

            processed += 1
            progress.progress(processed/total_files)
            status.text(f"معالجة الصور: {processed}/{total_files}")

        # 2) المطابقة بالأسماء بين القائمتين
        pairs = []          # [(unified_rec, voter_rec, score)]
        unmatched_unified = []
        unmatched_voter   = [v for v in cards["voter"]]  # سنزيل ما يُطابق لاحقًا

        # فهرس سريع للأسماء في بطاقات الناخب
        voter_names = [v[0] for v in cards["voter"]]

        for u in cards["unified"]:
            u_name_norm, u_name_raw, u_img = u
            if not voter_names:
                unmatched_unified.append(u)
                continue

            # أفضل تطابق
            match, score, idx = process.extractOne(
                u_name_norm,
                voter_names,
                scorer=fuzz.token_sort_ratio
            ) if voter_names else (None, 0, None)

            if match is not None and score >= 70:
                voter_rec = cards["voter"][idx]
                pairs.append((u, voter_rec, score))
                # إزالة هذا الناخب من قائمة غير المطابقة
                voter_names.pop(idx)
                unmatched_voter.pop(idx)
            else:
                unmatched_unified.append(u)

        # 3) إنشاء PDF: جدول عمودين (يسار: موحدة | يمين: ناخب)
        if pairs:
            pdf_name = "matched_cards.pdf"
            c = canvas.Canvas(pdf_name, pagesize=A4)
            page_w, page_h = A4
            margin = 36  # 0.5 inch تقريبًا
            col_gap = 24
            row_gap = 30
            img_max_w = (page_w - (2*margin) - col_gap) / 2.0
            img_max_h = 220  # ارتفاع مناسب للصورة
            text_gap = 14

            def _draw_label_center(x_center, y_text, txt):
                try:
                    c.setFont(arabic_font, 11)
                except:
                    c.setFont("Helvetica", 10)
                c.drawCentredString(x_center, y_text, fix_arabic_text(txt))

            row_y = page_h - margin
            col_left_x  = margin
            col_right_x = margin + img_max_w + col_gap

            items_in_page = 0
            max_rows_per_page = 3  # كل صف صورتين + نصوص

            for (u_norm, u_raw, u_img), (v_norm, v_raw, v_img), score in pairs:
                # سطر جديد؟
                if items_in_page >= max_rows_per_page:
                    c.showPage()
                    row_y = page_h - margin
                    items_in_page = 0

                # حساب قياسات الصورة المرسومة مع الحفاظ على النسبة
                def _draw_img(img_bgr, x, top_y):
                    ih, iw = img_bgr.shape[:2]
                    scale = min(img_max_w/iw, img_max_h/ih)
                    dw, dh = iw*scale, ih*scale
                    # تحويل إلى ملف مؤقت PNG
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    Image.fromarray(cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)).save(tmp.name)
                    c.drawImage(tmp.name, x, top_y - dh, width=dw, height=dh)
                    os.unlink(tmp.name)
                    return dw, dh

                # يسار: البطاقة الموحدة
                dw_u, dh_u = _draw_img(u_img, col_left_x, row_y)
                _draw_label_center(col_left_x + dw_u/2, row_y - dh_u - text_gap, u_raw if u_raw else "—")

                # يمين: بطاقة الناخب
                dw_v, dh_v = _draw_img(v_img, col_right_x, row_y)
                _draw_label_center(col_right_x + dw_v/2, row_y - dh_v - text_gap, v_raw if v_raw else "—")

                # التالي عموديًا
                row_y = row_y - max(dh_u, dh_v) - (text_gap + 26) - row_gap
                items_in_page += 1

            c.save()

            with open(pdf_name, "rb") as f:
                st.download_button("⬇️ تحميل PDF المطابقات", f, file_name=pdf_name, mime="application/pdf")

            st.success(f"✅ تم إنشاء ملف PDF للمطابقات: عدد الصفوف = {len(pairs)}")
        else:
            st.warning("⚠️ لم يتم العثور على أي أزواج متطابقة بالاسم (موحّدة ↔ ناخب).")

        # 4) عرض غير المطابق + المجهول
        if unmatched_unified or unmatched_voter or leftovers:
            st.markdown("### 🔎 بطاقات لم تُطابق / غير مُعرّفة")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("**البطاقات الوطنية غير المطابقة**")
                for nrm, raw, img in unmatched_unified[:6]:
                    st.image(cv2.cvtColor(img, cv2.COLOR_BGR2RGB), caption=raw or "—", use_container_width=True)
                if len(unmatched_unified) > 6:
                    st.info(f"+ {len(unmatched_unified)-6} أخرى...")
            with col2:
                st.markdown("**بطاقات الناخب غير المطابقة**")
                for nrm, raw, img in unmatched_voter[:6]:
                    st.image(cv2.cvtColor(img, cv2.COLOR_BGR2RGB), caption=raw or "—", use_container_width=True)
                if len(unmatched_voter) > 6:
                    st.info(f"+ {len(unmatched_voter)-6} أخرى...")
            with col3:
                st.markdown("**بطاقات غير مُعرّفة النوع/الاسم**")
                for nrm, raw, img in leftovers[:6]:
                    st.image(cv2.cvtColor(img, cv2.COLOR_BGR2RGB), caption=raw or "—", use_container_width=True)
                if len(leftovers) > 6:
                    st.info(f"+ {len(leftovers)-6} أخرى...")
