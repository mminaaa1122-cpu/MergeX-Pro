"""
Member Checker for MergeX Pro
-----------------------------
Reads the uploaded Clean Batch Excel file, looks up each mobile number in SQL Server,
classifies membership by the latest Product_Date record and Expiry_Date (MMYY),
then exports Members.xlsx and Old_Members_Not_Found.xlsx.

Configuration (recommended: Streamlit secrets):
[sqlserver]
server = "SERVER_IP_OR_HOST,1433"
database = "DATABASE_NAME"
username = "READ_ONLY_USERNAME"
password = "PASSWORD"
driver = "ODBC Driver 18 for SQL Server"
encrypt = true
trust_server_certificate = false

Install dependencies:
    pip install pyodbc openpyxl

Important:
- Use a read-only SQL login.
- Never commit credentials to Git.
- Confirm with your IT team whether ODBC Driver 18 and VPN/firewall access are configured.
"""

from __future__ import annotations

import os
import re
import io
from datetime import date, datetime
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from ops_log import log_operation


# The SQL query follows the schema supplied by the company.
# ROW_NUMBER selects the latest Product_Date per normalized mobile number.
# The query is read-only. It intentionally does not select unnecessary personal fields.
SQL_QUERY = r"""
SELECT
    c.[Card Holder Name] AS Card_Holder_Name,
    c.Mobile AS Mobile,
    c.[Expiry Date] AS Expiry_Date,
    PT.Product_Date AS Product_Date,
    PT.Delivery_Date AS Delivery_Date,
    PII.Product_Name AS Product_Name,
    PS.Product_Status AS Product_Status
FROM dbo.Tbl_Card AS c
LEFT JOIN dbo.tbl_Product_Transaction AS PT
    ON c.Card_ID = PT.Card_ID
LEFT JOIN dbo.Tbl_Product_Info AS PII
    ON PT.Product_Id = PII.Product_ID
LEFT JOIN dbo.Tbl_Product_Status AS PS
    ON PT.Product_Status_Id = PS.Product_Status_Id
WHERE c.Mobile IS NOT NULL;
"""


def normalize_mobile(value: Any) -> str:
    """Normalize formatting only; does not assume country-code equivalence."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"", "nan", "none", "nat"}:
        return ""
    # Excel may turn numeric-looking values into strings ending in .0
    text = re.sub(r"\.0$", "", text)
    # Keep digits and an initial + only, then compare digit sequence
    digits = re.sub(r"\D", "", text)
    return digits


def normalize_name(value: Any) -> str:
    """Normalize whitespace and case for a practical name comparison."""
    if pd.isna(value):
        return ""
    text = str(value).strip().casefold()
    return re.sub(r"\s+", " ", text)


def parse_expiry(value: Any) -> tuple[pd.Period | None, str]:
    """
    Parse company Expiry Date values such as '0625' as June 2025.
    Also accepts date/datetime objects and common full-date strings.
    Returns (month_period, display_value).
    """
    if pd.isna(value):
        return None, ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        ts = pd.Timestamp(value)
        return ts.to_period("M"), ts.strftime("%m/%Y")

    raw = str(value).strip()
    if not raw or raw.lower() in {"nan", "none", "nat"}:
        return None, ""

    # Common compact formats: MMYY and MMYYYY
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 4 and digits.isdigit():
        month, year = int(digits[:2]), 2000 + int(digits[2:])
        if 1 <= month <= 12:
            return pd.Period(year=year, month=month, freq="M"), f"{month:02d}/{year}"
    if len(digits) == 6 and digits.isdigit():
        month, year = int(digits[:2]), int(digits[2:])
        if 1 <= month <= 12 and 1900 <= year <= 2200:
            return pd.Period(year=year, month=month, freq="M"), f"{month:02d}/{year}"

    # Full date strings; day-first supports values such as 20/04/2024
    parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)
    if not pd.isna(parsed):
        return parsed.to_period("M"), parsed.strftime("%m/%Y")

    return None, raw


def get_sql_config() -> dict[str, Any]:
    """Read credentials from st.secrets first, then environment variables."""
    secret_config: dict[str, Any] = {}
    try:
        if "sqlserver" in st.secrets:
            secret_config = dict(st.secrets["sqlserver"])
    except Exception:
        secret_config = {}

    def val(key: str, env_key: str | None = None, default: Any = None) -> Any:
        if key in secret_config and secret_config[key] not in (None, ""):
            return secret_config[key]
        return os.getenv(env_key or f"MERGEX_SQL_{key.upper()}", default)

    config = {
        "server": val("server"),
        "database": val("database"),
        "username": val("username"),
        "password": val("password"),
        "driver": val("driver", default="ODBC Driver 18 for SQL Server"),
        "encrypt": val("encrypt", default=True),
        "trust_server_certificate": val("trust_server_certificate", default=False),
    }
    missing = [k for k in ("server", "database", "username", "password") if not config[k]]
    if missing:
        raise ValueError(
            "إعدادات SQL Server غير مكتملة. أضف القيم التالية إلى .streamlit/secrets.toml "
            "أو متغيرات البيئة: " + ", ".join(missing)
        )
    return config


def fetch_company_members() -> pd.DataFrame:
    """Connect to SQL Server and fetch the latest record per mobile."""
    try:
        import pyodbc
    except ImportError as exc:
        raise RuntimeError(
            "مكتبة pyodbc غير مثبتة. نفّذ: pip install pyodbc"
        ) from exc

    cfg = get_sql_config()
    encrypt = str(cfg["encrypt"]).lower() in {"1", "true", "yes", "on"}
    trust_cert = str(cfg["trust_server_certificate"]).lower() in {"1", "true", "yes", "on"}

    connection_string = (
        f"DRIVER={{{cfg['driver']}}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        f"Encrypt={'yes' if encrypt else 'no'};"
        f"TrustServerCertificate={'yes' if trust_cert else 'no'};"
        "Connection Timeout=15;"
    )

    connection = None
    try:
        connection = pyodbc.connect(connection_string, timeout=15)
        return pd.read_sql_query(SQL_QUERY, connection)
    
    except Exception as exc:
        import re

        error_text = str(exc)

        # Hide sensitive connection values before showing the error.
        for key in ("password", "username", "server", "database"):
            secret = str(cfg.get(key, "") or "")
            if secret:
                error_text = error_text.replace(secret, "[HIDDEN]")

        error_text = re.sub(
            r"(?i)(PWD|PASSWORD|UID|USER ID)\s*=\s*[^;]*",
            r"\1=[HIDDEN]",
            error_text,
        )

        raise RuntimeError(
            f"SQL Server diagnostic: {type(exc).__name__}: "
            f"{error_text[:1000]}"
        ) from None
    finally:
        if connection is not None:
            connection.close()



def prepare_company_records(company_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Keep all product transactions grouped by normalized mobile number."""
    required = {
        "Mobile", "Card_Holder_Name", "Expiry_Date", "Product_Date", "Product_Name"
    }
    missing = required - set(company_df.columns)
    if missing:
        raise ValueError("بيانات SQL ينقصها: " + ", ".join(sorted(missing)))

    df = company_df.copy()
    df["Mobile_Key"] = df["Mobile"].map(normalize_mobile)
    df["Name_Key"] = df["Card_Holder_Name"].map(normalize_name)
    df["Product_Date_Parsed"] = pd.to_datetime(
        df["Product_Date"], errors="coerce", dayfirst=True
    )
    df = df[df["Mobile_Key"] != ""].copy()
    df = df.sort_values("Product_Date_Parsed", ascending=True, na_position="first")
    return {str(mobile): group.copy() for mobile, group in df.groupby("Mobile_Key", sort=False)}


def classify_members(
    bank_df: pd.DataFrame,
    company_df: pd.DataFrame,
    as_of: date | None = None,
    mobile_column: str | None = None,
) -> pd.DataFrame:
    """Match mobile, flag name mismatches, classify product, and check expiry."""
    if bank_df.empty:
        raise ValueError("ملف العملاء فاضي.")

    if mobile_column is None:
        candidates = [
            c for c in bank_df.columns
            if str(c).strip().lower() in {
                "mobile", "mobile number", "phone", "phone number",
                "رقم الموبايل", "الموبايل", "موبايل"
            }
        ]
        if not candidates:
            raise ValueError("مش لاقي عمود الموبايل في ملف الباتش.")
        mobile_column = candidates[0]
    if mobile_column not in bank_df.columns:
        raise ValueError(f"عمود الموبايل '{mobile_column}' مش موجود في الملف.")

    name_candidates = [
        c for c in bank_df.columns
        if str(c).strip().casefold() in {
            "customer name", "customername", "name", "card holder name"
        }
    ]
    if not name_candidates:
        raise ValueError("مش لاقي عمود Customer Name في ملف الباتش.")
    name_column = name_candidates[0]

    company_by_mobile = prepare_company_records(company_df)
    current_period = pd.Timestamp(as_of or date.today()).to_period("M")
    output = bank_df.copy()
    for col in ["Database Name", "Name Match", "Product Type", "Product_Name",
                "Product_Date", "Expiry Date", "Product Status", "Status"]:
        output[col] = ""

    for idx, row in output.iterrows():
        mobile_key = normalize_mobile(row.get(mobile_column))
        batch_name = normalize_name(row.get(name_column))

        if not mobile_key or mobile_key not in company_by_mobile:
            output.at[idx, "Name Match"] = "No mobile match"
            output.at[idx, "Product Type"] = "Not Found"
            output.at[idx, "Status"] = "Not Found"
            continue

        all_records = company_by_mobile[mobile_key]
        matching_names = all_records[all_records["Name_Key"] == batch_name] if batch_name else all_records.iloc[0:0]
        if not matching_names.empty:
            # Exact name match: keep this field empty; only mismatches need a flag.
            output.at[idx, "Name Match"] = ""
            records = all_records.copy()
        else:
            # Name mismatch is a warning only: continue the check by mobile.
            output.at[idx, "Name Match"] = "Name Mismatch"
            records = all_records.copy()

        records = records.sort_values("Product_Date_Parsed", ascending=True, na_position="first")
        latest_record = records.iloc[-1]
        database_name = latest_record.get("Card_Holder_Name", "")
        output.at[idx, "Database Name"] = "" if pd.isna(database_name) else str(database_name)

        product_names = records["Product_Name"].fillna("").astype(str).str.strip()
        is_adib = product_names.str.casefold().eq("abu dhabi islamic bank".casefold())
        adib_records = records[is_adib]
        other_records = records[(product_names != "") & ~is_adib]

        if not adib_records.empty and not other_records.empty:
            output.at[idx, "Product Type"] = "ADIB + Xclusives"
        elif not adib_records.empty:
            output.at[idx, "Product Type"] = "ADIB"
        elif not other_records.empty:
            output.at[idx, "Product Type"] = "Xclusives"
        else:
            output.at[idx, "Product Type"] = "Product Empty"

        # Membership is checked against the latest non-ADIB transaction if available.
        selected_records = other_records if not other_records.empty else adib_records
        if selected_records.empty:
            selected_records = records
        selected_records = selected_records.sort_values(
            "Product_Date_Parsed", ascending=True, na_position="first"
        )
        selected = selected_records.iloc[-1]

        product_name = selected.get("Product_Name", "")
        product_date = pd.to_datetime(selected.get("Product_Date"), errors="coerce", dayfirst=True)
        expiry_period, expiry_display = parse_expiry(selected.get("Expiry_Date"))
        product_status = selected.get("Product_Status", "")

        output.at[idx, "Product_Name"] = "" if pd.isna(product_name) else str(product_name)
        output.at[idx, "Product_Date"] = "" if pd.isna(product_date) else product_date.strftime("%d/%m/%Y")
        output.at[idx, "Expiry Date"] = expiry_display
        output.at[idx, "Product Status"] = "" if pd.isna(product_status) else str(product_status)

        # ADIB always counts as a Member, regardless of expiry date.
        if not adib_records.empty:
            output.at[idx, "Status"] = "Member"
        elif expiry_period is None:
            output.at[idx, "Status"] = "Needs Review"
        elif expiry_period < current_period:
            output.at[idx, "Status"] = "Old Member"
        else:
            output.at[idx, "Status"] = "Member"

    return output


def export_member_reports(results: pd.DataFrame, output_dir: str = ".") -> dict[str, Any]:
    """Create formatted Members, Old Members/Not Found, ADIB, and Xclusives reports."""
    from pathlib import Path
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    product_types = results["Product Type"].fillna("").astype(str)
    members = results[results["Status"] == "Member"].copy()
    members["Status"] = members["Product Type"].map(
        lambda value: "ADIB" if "ADIB" in str(value)
        else ("Xclusives" if "Xclusives" in str(value) else "Product Empty")
    )
    others = results[results["Status"].isin(["Old Member", "Not Found", "Needs Review"])].copy()
    adib = results[product_types.str.contains("ADIB", case=False, na=False)].copy()
    xclusives = results[product_types.str.contains("Xclusives", case=False, na=False)].copy()
    paths = {
        "members_file": out / "Members.xlsx",
        "other_file": out / "Old_Members_Not_Found.xlsx",
        "adib_file": out / "ADIB_Customers.xlsx",
        "xclusives_file": out / "Xclusives_Customers.xlsx",
    }
    members.to_excel(paths["members_file"], index=False)
    others.to_excel(paths["other_file"], index=False)
    adib.to_excel(paths["adib_file"], index=False)
    xclusives.to_excel(paths["xclusives_file"], index=False)
    colors = {"ADIB": ("FFF2CC", "7F6000"), "Xclusives": ("E2EFDA", "375623"), "Old Member": ("D9EAF7", "1F4E78"), "Product Empty": ("EADCF8", "4B2E83")}
    for path in paths.values():
        wb = load_workbook(path)
        ws = wb.active
        ws.freeze_panes = "A2"
        if ws.max_row and ws.max_column:
            ws.auto_filter.ref = ws.dimensions
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=0)
            ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(max_len + 2, 12), 42)
        header_map = {str(cell.value): cell.column for cell in ws[1]}
        status_index = header_map.get("Status")
        if status_index:
            for row_num in range(2, ws.max_row + 1):
                status = ws.cell(row_num, status_index).value
                if status in colors:
                    fill_hex, font_hex = colors[status]
                    fill = PatternFill(fill_type="solid", fgColor=fill_hex)
                    font = Font(color=font_hex)
                    for cell in ws[row_num]:
                        cell.fill, cell.font = fill, font
        wb.save(path)
    counts = results["Status"].value_counts().to_dict()
    return {
        **{key: str(value) for key, value in paths.items()},
        "members_count": int(counts.get("Member", 0)),
        "old_members_count": int(counts.get("Old Member", 0)),
        "not_found_count": int(counts.get("Not Found", 0)),
        "needs_review_count": int(counts.get("Needs Review", 0)),
        "adib_count": int(product_types.str.contains("ADIB", case=False, na=False).sum()),
        "empty_product_count": int((product_types == "Product Empty").sum()),
        "xclusives_count": int(product_types.str.contains("Xclusives", case=False, na=False).sum()),
        "name_mismatch_count": int((results["Name Match"] == "Name Mismatch").sum()),
        "total_count": int(len(results)),
    }

def render_member_checker() -> None:
    """Streamlit UI integrated into MergeX Pro."""
    st.markdown("<h1 style='text-align:center'>🔎 Member Checker</h1>", unsafe_allow_html=True)
    st.markdown(
        "<p style='text-align:center'>افحص ملف الـ Clean Batch مع بيانات الشركة.</p>",
        unsafe_allow_html=True,
    )
    st.divider()

    uploaded = st.file_uploader(
        "📂 ارفع ملف العملاء (Clean Batch)",
        type=["xlsx", "xls"],
        key="member_checker_bank_upload",
        help="لازم يحتوي على Customer Name وMobile.",
    )
    

    if st.button("🚀 Check Memberships", type="primary", use_container_width=True, key="member_checker_start"):
        if uploaded is None:
            st.warning("ارفع ملف العملاء الأول.")
            return
        try:
            bank_df = pd.read_excel(uploaded, dtype=str)
            with st.spinner("جاري الاتصال بقاعدة بيانات الشركة وفحص البيانات..."):
                company_df = fetch_company_members()
                results = classify_members(bank_df, company_df)
                report = export_member_reports(results, output_dir=os.getcwd())

            st.session_state["member_checker_results"] = results
            st.session_state["member_checker_report"] = report
            with open(report["members_file"], "rb") as f:
                st.session_state["member_checker_members_bytes"] = f.read()
            with open(report["other_file"], "rb") as f:
                st.session_state["member_checker_others_bytes"] = f.read()
            with open(report["adib_file"], "rb") as f:
                st.session_state["member_checker_adib_bytes"] = f.read()
            with open(report["xclusives_file"], "rb") as f:
                st.session_state["member_checker_xclusives_bytes"] = f.read()
            st.success("اكتمل الفحص بنجاح.")
            log_operation(
                "Member Checker", "فحص عضويات",
                file_name=getattr(uploaded, "name", ""),
                metrics={
                    "total": report["total_count"],
                    "members": report["members_count"],
                    "old_members": report["old_members_count"],
                    "not_found": report["not_found_count"],
                    "needs_review": report["needs_review_count"],
                    "adib": report["adib_count"],
                    "xclusives": report["xclusives_count"],
                    "empty_product": report["empty_product_count"],
                    "name_mismatch": report["name_mismatch_count"],
                },
            )
        except Exception as exc:
            st.error(str(exc))

    results = st.session_state.get("member_checker_results")
    report = st.session_state.get("member_checker_report")
    if results is not None and report is not None:
        
        # ===== Premium Summary & Downloads UI =====
        st.markdown("""
        <style>
        .mc-title {
            text-align: center;
            font-size: 25px;
            font-weight: 800;
            margin: 8px 0 22px;
            color: #f1f5f9;
        }
        .mc-card {
            background: linear-gradient(135deg, #202b45, #151d31);
            border: 1px solid #354362;
            border-radius: 15px;
            padding: 17px 18px;
            min-height: 105px;
            margin-bottom: 12px;
            box-shadow: 0 5px 16px rgba(0,0,0,.12);
        }
        .mc-label {
            color: #b5c3df;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 10px;
        }
        .mc-value {
            color: #ffffff;
            font-size: 29px;
            font-weight: 800;
            line-height: 1.2;
        }
        .mc-total { border-top: 4px solid #8b9cff; }
        .mc-name  { border-top: 4px solid #f6bd60; }
        .mc-adib  { border-top: 4px solid #48d6a3; }
        .mc-xcl   { border-top: 4px solid #55b8ff; }
        .mc-member{ border-top: 4px solid #4ade80; }
        .mc-old   { border-top: 4px solid #fb923c; }
        .mc-not   { border-top: 4px solid #f87171; }
        .mc-empty { border-top: 4px solid #a78bfa; }
        div.stDownloadButton > button {
            min-height: 52px;
            border-radius: 12px;
            border: 1px solid #3b4967;
            font-weight: 700;
            transition: all .18s ease;
        }
        div.stDownloadButton > button:hover {
            border-color: #8b9cff;
            transform: translateY(-2px);
        }
        </style>
        <div class="mc-title">📊 Test Summary</div>
        """, unsafe_allow_html=True)

        def mc_card(label, value, style):
            st.markdown(
                f'<div class="mc-card {style}">'
                f'<div class="mc-label">{label}</div>'
                f'<div class="mc-value">{value}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        r1a, r1b, r1c, r1d = st.columns(4)
        with r1a:
            mc_card("إجمالي الباتش", report["total_count"], "mc-total")
        with r1b:
            mc_card("حالات اختلاف الاسم", report["name_mismatch_count"], "mc-name")
        with r1c:
            mc_card("ADIB", report["adib_count"], "mc-adib")

        with r1d:
            mc_card("Product Name فاضي", report["empty_product_count"], "mc-empty")

        r2a, r2b, r2c, r2d = st.columns(4)
        with r2a:
            mc_card("Xclusives", report["xclusives_count"], "mc-xcl")
        with r2b:
            mc_card("Members", report["members_count"], "mc-member")
        with r2c:
            mc_card("Old Members", report["old_members_count"], "mc-old")
        with r2d:
            mc_card("Not Found", report["not_found_count"], "mc-not")

        if report["needs_review_count"]:
            st.info(
                f"حالات تحتاج مراجعة بسبب تاريخ انتهاء غير واضح: "
                f"{report['needs_review_count']}"
            )

        st.markdown("### 📥 Download the Reports")

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "🟢  تحميل تقرير ADIB",
                data=st.session_state["member_checker_adib_bytes"],
                file_name="ADIB_Customers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.download_button(
                "📊  تحميل تقرير Members",
                data=st.session_state["member_checker_members_bytes"],
                file_name="Members.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with d2:
            st.download_button(
                "🔵  تحميل تقرير Xclusives",
                data=st.session_state["member_checker_xclusives_bytes"],
                file_name="Xclusives_Customers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.download_button(
                "🟠  تحميل Old Members + Not Found",
                data=st.session_state["member_checker_others_bytes"],
                file_name="Old_Members_Not_Found.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
