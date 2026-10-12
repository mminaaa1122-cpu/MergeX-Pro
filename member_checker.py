
from pathlib import Path
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

import streamlit as st
import tempfile


# =========================================================
# Member Checker - إعدادات أساسية
# =========================================================

MOBILE_COLUMN = "Mobile"
EXPIRY_COLUMN = "Expiry Date"

MEMBERS_FILE = "Members.xlsx"
OLD_MEMBERS_FILE = "Old_Members_Not_Found.xlsx"


def normalize_mobile(value):
    """
    توحيد شكل رقم الموبايل للمقارنة فقط.
    لا يغيّر الرقم الأصلي الموجود في ملفات Excel.
    """
    if pd.isna(value):
        return ""

    mobile = str(value).strip()

    # معالجة الأرقام التي يقرأها Excel مثل 1012345678.0
    mobile = re.sub(r"\.0$", "", mobile)

    # إزالة المسافات والشرطات والأقواس
    mobile = re.sub(r"[\s\-\(\)]", "", mobile)

    return mobile


def prepare_company_records(company_df):
    """
    تجهيز سجلات الشركة واستخراج أحدث Expiry Date لكل Mobile.
    لا يعتمد على Delivery Date لتحديد حالة العضوية.
    """
    required = [MOBILE_COLUMN, EXPIRY_COLUMN]
    missing = [col for col in required if col not in company_df.columns]

    if missing:
        raise ValueError(
            "أعمدة بيانات الشركة المطلوبة غير موجودة: "
            + ", ".join(missing)
        )

    records = company_df.copy()

    records["_mobile_key"] = records[MOBILE_COLUMN].apply(
        normalize_mobile
    )

    # تحويل تاريخ انتهاء العضوية إلى تاريخ قابل للمقارنة.
    records["_expiry_parsed"] = pd.to_datetime(
        records[EXPIRY_COLUMN],
        errors="coerce",
        dayfirst=True
    )

    # تجاهل السجلات التي لا تحتوي على رقم موبايل.
    records = records[records["_mobile_key"] != ""].copy()

    # نحتفظ بأحدث تاريخ صالح لكل رقم.
    latest_expiry = (
        records.dropna(subset=["_expiry_parsed"])
        .groupby("_mobile_key")["_expiry_parsed"]
        .max()
        .to_dict()
    )

    # نحتفظ أيضًا بالأرقام الموجودة حتى لو لم يكن لها تاريخ صالح.
    known_mobiles = set(records["_mobile_key"])

    return latest_expiry, known_mobiles


def classify_members(bank_df, company_df, as_of=None):
    """
    يقارن أرقام ملف البنك بسجلات الشركة.

    Member:
        تاريخ الانتهاء في الشهر الحالي أو في المستقبل.

    Old Member:
        تاريخ الانتهاء قبل الشهر الحالي.

    Not Found:
        رقم الموبايل غير موجود في سجلات الشركة.

    Needs Review:
        الرقم موجود لكن لا يوجد له تاريخ انتهاء صالح.
        لا نصنفه كعضو قديم بدون دليل.
    """
    if MOBILE_COLUMN not in bank_df.columns:
        raise ValueError(
            f"ملف البنك يجب أن يحتوي على عمود {MOBILE_COLUMN}"
        )

    if as_of is None:
        as_of = pd.Timestamp.today().normalize()
    else:
        as_of = pd.Timestamp(as_of)

    current_month = as_of.to_period("M")

    latest_expiry, known_mobiles = prepare_company_records(company_df)

    results = bank_df.copy()

    results["_mobile_key"] = results[MOBILE_COLUMN].apply(
        normalize_mobile
    )

    statuses = []
    expiry_dates = []

    for mobile in results["_mobile_key"]:
        if not mobile or mobile not in known_mobiles:
            statuses.append("Not Found")
            expiry_dates.append(pd.NaT)
            continue

        expiry = latest_expiry.get(mobile)

        if expiry is None or pd.isna(expiry):
            statuses.append("Needs Review")
            expiry_dates.append(pd.NaT)
            continue

        expiry_dates.append(expiry)

        if expiry.to_period("M") >= current_month:
            statuses.append("Member")
        else:
            statuses.append("Old Member")

    results["Expiry Date"] = expiry_dates
    results["Status"] = statuses

    # عمود المقارنة داخلي فقط، لا يظهر في الملفات النهائية.
    results.drop(columns=["_mobile_key"], inplace=True)

    return results


def style_old_member_rows(file_path):
    """
    تلوين صفوف Old Member باللون الأزرق الفاتح.
    """
    workbook = load_workbook(file_path)

    blue_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7"
    )

    blue_font = Font(color="1F4E78")

    for worksheet in workbook.worksheets:
        headers = {
            cell.value: cell.column
            for cell in worksheet[1]
        }

        status_col = headers.get("Status")

        if status_col is None:
            continue

        for row_num in range(2, worksheet.max_row + 1):
            status = worksheet.cell(
                row=row_num,
                column=status_col
            ).value

            if status == "Old Member":
                for cell in worksheet[row_num]:
                    cell.fill = blue_fill
                    cell.font = blue_font

    workbook.save(file_path)


def export_member_reports(results, output_dir="."):
    """
    إنشاء الملفين:
    Members.xlsx
    Old_Members_Not_Found.xlsx
    """
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    members = results[
        results["Status"] == "Member"
    ].copy()

    other_records = results[
        results["Status"] != "Member"
    ].copy()

    members_path = output_path / MEMBERS_FILE
    other_path = output_path / OLD_MEMBERS_FILE

    members.to_excel(members_path, index=False)
    other_records.to_excel(other_path, index=False)

    style_old_member_rows(other_path)

    return {
        "members_file": str(members_path),
        "other_file": str(other_path),
        "members_count": len(members),
        "old_members_count": int(
            (other_records["Status"] == "Old Member").sum()
        ),
        "not_found_count": int(
            (other_records["Status"] == "Not Found").sum()
        ),
        "needs_review_count": int(
            (other_records["Status"] == "Needs Review").sum()
        ),
    }


def process_members(bank_df, company_df, output_dir="."):
    """
    نقطة التشغيل الرئيسية التي سيستدعيها التطبيق لاحقًا.
    """
    results = classify_members(
        bank_df=bank_df,
        company_df=company_df
    )

    report = export_member_reports(
        results=results,
        output_dir=output_dir
    )

    return results, report




def render_member_checker():
    st.markdown(
        "<h1>🔎 Member Checker</h1>",
        unsafe_allow_html=True
    )

    st.markdown(
        "تحقق من حالة العضويات باستخدام رقم الموبايل "
        "وأحدث تاريخ انتهاء للعضوية."
    )

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📂 ملف العملاء")
        bank_file = st.file_uploader(
            "ارفع ملف Clean Batch",
            type=["xlsx"],
            key="member_bank_upload",
            help="الملف الذي يحتوي على عمود Mobile."
        )

    with col2:
        st.markdown("### 🗄️ بيانات العضويات")
        company_file = st.file_uploader(
            "ملف تجريبي من بيانات الشركة",
            type=["xlsx"],
            key="member_company_upload",
            help="للاختبار المؤقت فقط: يجب أن يحتوي على Mobile و Expiry Date."
        )

    st.markdown("---")

    st.markdown("### 📊 ملخص الفحص")

    c1, c2, c3 = st.columns(3)
    c1.metric("Members", "—")
    c2.metric("Old Members", "—")
    c3.metric("Not Found", "—")

    st.info(
        "الاتصال المباشر بنظام الشركة لم يتم إعداده بعد. "
        "يمكنك تجربة الواجهة باستخدام ملف Excel تجريبي "
        "بدلًا من الاتصال بالنظام."
    )

    if st.button(
        "🚀 بدء فحص العضويات",
        key="run_member_check",
        use_container_width=True
    ):
        if bank_file is None or company_file is None:
            st.warning(
                "ارفع ملف Clean Batch وملف بيانات الشركة التجريبي أولًا."
            )
            return

        try:
            bank_df = pd.read_excel(bank_file, dtype=str)
            company_df = pd.read_excel(company_file, dtype=str)

            with st.spinner("جاري فحص العضويات..."):
                results = classify_members(
                    bank_df=bank_df,
                    company_df=company_df
                )

            with tempfile.TemporaryDirectory() as temp_dir:
                report = export_member_reports(
                    results,
                    output_dir=temp_dir
                )

                members_bytes = Path(
                    report["members_file"]
                ).read_bytes()

                other_bytes = Path(
                    report["other_file"]
                ).read_bytes()

            st.session_state["member_results"] = results
            st.session_state["member_report_files"] = {
                "members": members_bytes,
                "other": other_bytes,
            }
            st.session_state["member_report_counts"] = report

        except Exception as e:
            st.error(f"تعذر إكمال الفحص: {e}")

    results = st.session_state.get("member_results")

    if results is not None:
        report = st.session_state["member_report_counts"]
        files = st.session_state["member_report_files"]

        st.markdown("### ✅ نتائج الفحص")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Members", report["members_count"])
        c2.metric("Old Members", report["old_members_count"])
        c3.metric("Not Found", report["not_found_count"])
        c4.metric("Needs Review", report["needs_review_count"])

        st.dataframe(
            results.head(100),
            use_container_width=True,
            hide_index=True
        )

        d1, d2 = st.columns(2)

        with d1:
            st.download_button(
                "📥 تحميل Members.xlsx",
                data=files["members"],
                file_name="Members.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        with d2:
            st.download_button(
                "📥 تحميل Old Members & Not Found",
                data=files["other"],
                file_name="Old_Members_Not_Found.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )