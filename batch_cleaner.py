import streamlit as st
import pandas as pd
from io import BytesIO
import os
from datetime import datetime
import pytz

# ==================================================
# إعدادات
# ==================================================
MASTER_FILE = "master_database.parquet"
LOG_FILE = "master_update_log.csv"
SEARCH_COLUMN = "Mobile"

# ←←← حط هنا اليوزرين اللي ليهم صلاحية إدارة البيانات
MASTER_ALLOWED_USERS = ["admin", "salma"]   # غيّر الاسم الثاني زي ما تحب


# ==================================================
# دوال مساعدة
# ==================================================
def normalize_mobile(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"[\s\-\(\)]", "", regex=True)
    )


def get_egypt_time():
    return datetime.now(pytz.timezone("Africa/Cairo")).strftime("%Y-%m-%d %H:%M:%S")


def create_excel_download(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#4F46E5",
            "font_color": "#FFFFFF",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        })

        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_format)

        for col_num, col_name in enumerate(df.columns):
            max_len = len(str(col_name))
            if not df.empty:
                sample = df[col_name].astype(str).head(400)
                for v in sample:
                    max_len = max(max_len, len(str(v)))
            worksheet.set_column(col_num, col_num, min(max(max_len + 2, 12), 40))

        worksheet.freeze_panes(1, 0)
        if len(df.columns) > 0:
            worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

    output.seek(0)
    return output.getvalue()


def load_master() -> pd.DataFrame:
    if not os.path.exists(MASTER_FILE):
        return pd.DataFrame()
    try:
        return pd.read_parquet(MASTER_FILE)
    except Exception:
        return pd.DataFrame()


def save_master(df: pd.DataFrame):
    df.to_parquet(MASTER_FILE, index=False)


def get_master_mobile_set() -> set:
    df = load_master()
    if df.empty or SEARCH_COLUMN not in df.columns:
        return set()
    mobiles = normalize_mobile(df[SEARCH_COLUMN])
    return set(mobiles[mobiles != ""].tolist())


def log_update(action: str, before: int, after: int, added: int, username: str):
    # الصف الجديد فقط
    new_row = {
        "الوقت": get_egypt_time(),
        "المستخدم": username,
        "العملية": action,
        "قبل": before,
        "بعد": after,
        "المضاف": added,
    }
    new_df = pd.DataFrame([new_row])

    # 1. الحفظ في CSV
    if os.path.exists(LOG_FILE):
        old_csv = pd.read_csv(LOG_FILE)
        full_csv = pd.concat([old_csv, new_df], ignore_index=True)
    else:
        full_csv = new_df

    full_csv.to_csv(LOG_FILE, index=False)

    # 2. النسخ الاحتياطي في Excel
    BACKUP_LOG_FILE = "master_update_log_backup.xlsx"
    try:
        if os.path.exists(BACKUP_LOG_FILE):
            old_backup = pd.read_excel(BACKUP_LOG_FILE, dtype=str)
            full_backup = pd.concat([old_backup, new_df], ignore_index=True)
        else:
            full_backup = new_df

        full_backup.to_excel(BACKUP_LOG_FILE, index=False)
    except Exception as e:
        print(f"خطأ في حفظ الـ Backup: {e}")


# ==================================================
# قسم إدارة البيانات (للمسؤولين فقط)
# ==================================================
def render_master_management():
    st.markdown("## 🔐 إدارة البيانات الرئيسية")
    st.caption("هذا القسم متاح فقط للمستخدمين المصرح لهم")

    master_df = load_master()
    master_count = len(master_df)

    unique_count = 0
    duplicate_count = 0

    if master_count > 0 and SEARCH_COLUMN in master_df.columns:
        mobiles = normalize_mobile(master_df[SEARCH_COLUMN])
        mobiles = mobiles[mobiles != ""]
        unique_count = mobiles.nunique()
        duplicate_count = len(mobiles) - unique_count

    # ---- البطاقات ----
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #4f46e5, #7c3aed); padding: 1.2rem; border-radius: 16px; text-align: center;">
            <p style="color: #c4b5fd; margin: 0; font-size: 0.9rem;">إجمالي السجلات</p>
            <p style="color: white; font-size: 1.8rem; font-weight: 800; margin: 0.3rem 0 0 0;">{master_count:,}</p>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #059669, #10b981); padding: 1.2rem; border-radius: 16px; text-align: center;">
            <p style="color: #a7f3d0; margin: 0; font-size: 0.9rem;">موبايلات فريدة</p>
            <p style="color: white; font-size: 1.8rem; font-weight: 800; margin: 0.3rem 0 0 0;">{unique_count:,}</p>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        last_update = "—"
        if os.path.exists(LOG_FILE):
            try:
                log = pd.read_csv(LOG_FILE)
                if not log.empty:
                    last_update = str(log.iloc[-1]["الوقت"])[:16]
            except Exception:
                pass
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #d97706, #f59e0b); padding: 1.2rem; border-radius: 16px; text-align: center;">
            <p style="color: #fde68a; margin: 0; font-size: 0.9rem;">آخر تحديث</p>
            <p style="color: white; font-size: 1.1rem; font-weight: 700; margin: 0.3rem 0 0 0;">{last_update}</p>
        </div>
        """, unsafe_allow_html=True)

    # ---- رسالة المكررين ----
        # ---- رسالة المكررين ----
    if duplicate_count > 0:
        # عدد الصفوف اللي فيها تكرار (اللي بتنزل في ملف التصدير)
        mobiles = normalize_mobile(master_df[SEARCH_COLUMN])
        rows_with_dup = int((mobiles.duplicated(keep=False) & (mobiles != "")).sum())

        st.markdown(f"""
        <div style="background: rgba(239, 68, 68, 0.1); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: 12px; padding: 0.9rem 1.2rem; margin-top: 1rem; text-align: center;">
            <p style="color: #fca5a5; margin: 0 0 0.4rem 0; font-size: 1.05rem;">
                ⚠️ في <b>{rows_with_dup:,}</b> صف رقم الموبايل بتاعه مكرر
            </p>
            <p style="color: #f87171; margin: 0; font-size: 0.92rem;">
                (عدد التكرارات الزيادة = {duplicate_count:,} &nbsp;|&nbsp; 84119 - 83628 = 491)
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.success("✅ مفيش أي أرقام موبايل مكررة في القاعدة")

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- أزرار التصدير تحت البطاقات ----
    if master_count > 0:
        e1, e2, e3 = st.columns(3)

        with e1:
            excel_all = create_excel_download(master_df, "All_Records")
            st.download_button(
                "📥 تصدير كل السجلات",
                data=excel_all,
                file_name=f"All_Records_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="export_all"
            )

        with e2:
            if SEARCH_COLUMN in master_df.columns:
                unique_df = master_df.drop_duplicates(subset=[SEARCH_COLUMN], keep="first")
                excel_unique = create_excel_download(unique_df, "Unique_Mobiles")
                st.download_button(
                    "📥 تصدير الموبايلات الفريدة",
                    data=excel_unique,
                    file_name=f"Unique_Mobiles_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="export_unique"
                )

        with e3:
            if SEARCH_COLUMN in master_df.columns and duplicate_count > 0:
                mobiles = normalize_mobile(master_df[SEARCH_COLUMN])
                dup_mask = mobiles.duplicated(keep=False) & (mobiles != "")
                dup_df = master_df.loc[dup_mask].copy()
                excel_dups = create_excel_download(dup_df, "Duplicate_Mobiles")
                st.download_button(
                    "📥 تصدير المكررين",
                    data=excel_dups,
                    file_name=f"Duplicate_Mobiles_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="export_dups"
                )
            else:
                st.button("📥 تصدير المكررين", disabled=True, use_container_width=True)

    st.markdown("---")

    # ---- إضافة بيانات ----
    st.markdown("### 🚀 إضافة بيانات جديدة")

    update_file = st.file_uploader(
        "ارفع ملف Excel (البيانات التاريخية أو السجلات الجديدة)",
        type=["xlsx", "xls"],
        key="master_upload_secure",
        help="بعد أول مرة يُفضل رفع السجلات الجديدة فقط"
    )

    if update_file:
        update_file.seek(0)
        df_new = pd.read_excel(update_file, dtype=str)
        df_new.columns = df_new.columns.astype(str).str.strip()

        if SEARCH_COLUMN not in df_new.columns:
            st.error("❌ الملف لازم يحتوي على عمود اسمه **Mobile**")
        else:
            existing_set = get_master_mobile_set()
            new_mobiles = normalize_mobile(df_new[SEARCH_COLUMN])
            is_new = (~new_mobiles.isin(existing_set)) & (new_mobiles != "")
            new_count = int(is_new.sum())
            dup_count = int(((new_mobiles.isin(existing_set)) & (new_mobiles != "")).sum())
            empty_count = int((new_mobiles == "").sum())

            st.markdown(f"""
            <div style="background: rgba(79, 70, 229, 0.1); border: 1px solid rgba(79, 70, 229, 0.3); border-radius: 12px; padding: 1rem; margin: 1rem 0;">
                <p style="margin: 0.2rem 0;">📊 إجمالي الصفوف في الملف: <b>{len(df_new):,}</b></p>
                <p style="margin: 0.2rem 0; color: #34d399;">✅ سجلات جديدة هتتضاف: <b>{new_count:,}</b></p>
                <p style="margin: 0.2rem 0; color: #fbbf24;">⚠️ موجودة مسبقاً (هتتجاهل): <b>{dup_count:,}</b></p>
                <p style="margin: 0.2rem 0; color: #f87171;">📭 موبايل فاضي: <b>{empty_count:,}</b></p>
            </div>
            """, unsafe_allow_html=True)

            if new_count > 0:
                if st.button("✅ تأكيد الإضافة الآن", type="primary", use_container_width=True):
                    before = master_count
                    df_to_add = df_new.loc[is_new].copy()

                    if master_df.empty:
                        final_df = df_to_add
                    else:
                        final_df = pd.concat([master_df, df_to_add], ignore_index=True, sort=False)

                    save_master(final_df)
                    after = len(final_df)

                    username = st.session_state.get("username", "unknown")
                    log_update("إضافة سجلات", before, after, new_count, username)

                    st.success(f"🎉 تم إضافة **{new_count:,}** سجل بنجاح! الإجمالي الآن **{after:,}**")
                    st.balloons()
                    st.rerun()
            else:
                st.warning("مفيش سجلات جديدة للإضافة في الملف ده.")

    st.markdown("---")

    # ---- أزرار التحكم ----
    
    if master_count > 0:
            if st.button("🗑️ مسح القاعدة بالكامل", use_container_width=True):
                st.session_state["confirm_clear_master"] = True

    if st.session_state.get("confirm_clear_master"):
        st.error("⚠️ تحذير: هل أنت متأكد من مسح كل البيانات؟ لا يمكن التراجع.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("نعم، امسح كل شيء", type="primary"):
                before = master_count
                if os.path.exists(MASTER_FILE):
                    os.remove(MASTER_FILE)
                log_update("مسح كامل", before, 0, 0, st.session_state.get("username", "unknown"))
                st.session_state["confirm_clear_master"] = False
                st.success("تم مسح القاعدة بالكامل")
                st.rerun()
        with c2:
            if st.button("إلغاء"):
                st.session_state["confirm_clear_master"] = False
                st.rerun()

    # ---- سجل العمليات ----
    st.markdown("---")
    st.markdown("### 📜 سجل العمليات")

    if os.path.exists(LOG_FILE):
        try:
            log_df = pd.read_csv(LOG_FILE)
            st.dataframe(log_df.tail(15), use_container_width=True, height=280)
        except Exception:
            st.info("تعذر قراءة السجل")
    else:
        st.info("لا توجد عمليات مسجلة بعد")


# ==================================================
# قسم فحص شيت البنك (متاح للجميع)
# ==================================================
def render_bank_check():
    st.markdown("## 📥 فحص شيت البنك")

    master_count = len(load_master())
    if master_count == 0:
        st.warning("⚠️ قاعدة البيانات الرئيسية فارغة. يجب على المسؤول رفع البيانات أولاً.")
        return

    st.caption(f"البحث يتم ضد **{master_count:,}** سجل • المفتاح: **Mobile**")

    bank_file = st.file_uploader(
        "ارفع ملف البنك هنا",
        type=["xlsx", "xls"],
        key="bank_uploader_main",
        help="لازم يحتوي على عمود Mobile"
    )

    if bank_file:
        try:
            bank_file.seek(0)
            preview = pd.read_excel(bank_file, dtype=str, nrows=2)
            preview.columns = preview.columns.astype(str).str.strip()
            if "Mobile" not in preview.columns:
                st.error("❌ ملف البنك لازم يحتوي على عمود **Mobile**")
            else:
                st.success("✅ تم التعرف على الملف بنجاح")
        except Exception as e:
            st.error(f"خطأ: {e}")

    if st.button("🔍 ابدأ الفحص الآن", type="primary", use_container_width=True, disabled=bank_file is None):
        with st.spinner("جاري البحث في القاعدة..."):
            try:
                bank_file.seek(0)
                bank_df = pd.read_excel(bank_file, dtype=str)
                bank_df.columns = bank_df.columns.astype(str).str.strip()

                if "Mobile" not in bank_df.columns:
                    st.error("عمود Mobile غير موجود")
                    return

                original_cols = bank_df.columns.tolist()
                bank_key = normalize_mobile(bank_df["Mobile"])
                master_set = get_master_mobile_set()

                dup_mask = bank_key.isin(master_set)

                dup_df = bank_df.loc[dup_mask, original_cols].copy().reset_index(drop=True)
                clean_df = bank_df.loc[~dup_mask, original_cols].copy().reset_index(drop=True)

                st.session_state["bc_clean"] = clean_df
                st.session_state["bc_dup"] = dup_df
                st.session_state["bc_summary"] = {
                    "total": len(bank_df),
                    "duplicates": len(dup_df),
                    "clean": len(clean_df),
                }
                st.success("✅ الفحص خلص بنجاح")
                st.balloons()

            except Exception as e:
                st.error(f"حدث خطأ: {e}")
                return

    if "bc_clean" not in st.session_state:
        return

    summary = st.session_state["bc_summary"]
    clean_df = st.session_state["bc_clean"]
    dup_df = st.session_state["bc_dup"]

    st.markdown("---")
    st.markdown("### 📊 النتائج")

    m1, m2, m3 = st.columns(3)
    m1.metric("إجمالي البنك", f"{summary['total']:,}")
    m2.metric("Duplicates", f"{summary['duplicates']:,}")
    m3.metric("Clean", f"{summary['clean']:,}")

    st.markdown("---")
    st.markdown("### 🟥 Duplicate Records")

    if dup_df.empty:
        st.success("🎉 مفيش أي تكرار")
    else:
        st.warning(f"تم العثور على **{len(dup_df):,}** سجل مكرر")
        with st.expander("👁️ معاينة Duplicates"):
            st.dataframe(dup_df.head(80), use_container_width=True, height=320)
        st.download_button(
            "📥 تحميل Duplicates",
            data=create_excel_download(dup_df, "Duplicates"),
            file_name="Duplicates.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_dup_btn",
        )

    st.markdown("---")
    st.markdown("### 🟢 Clean Batch")
    st.success(f"الملف النظيف فيه **{len(clean_df):,}** سجل")

    with st.expander("👁️ معاينة Clean"):
        st.dataframe(clean_df.head(80), use_container_width=True, height=320)

    st.download_button(
        "📥 تحميل Clean Batch",
        data=create_excel_download(clean_df, "Clean_Batch"),
        file_name="Clean_Batch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="dl_clean_btn",
    )


# ==================================================
# الواجهة الرئيسية
# ==================================================
def render_batch_cleaner():
    st.markdown("<h1>🧹 Batch Cleaner</h1>", unsafe_allow_html=True)
    st.markdown(
        """
        <p style="text-align:center; color:#94a3b8; font-size:1.05rem;">
        نظام تنظيف شيتات البنك • بيانات مخزنة بشكل دائم وآمن
        </p>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("---")

    current_user = st.session_state.get("username", "")
    is_allowed = current_user in MASTER_ALLOWED_USERS

    # ---- التبويبات ----
    if is_allowed:
        tab1, tab2 = st.tabs(["📥 فحص شيت البنك", "🔐 إدارة البيانات الرئيسية"])
        with tab1:
            render_bank_check()
        with tab2:
            render_master_management()
    else:
        # اليوزر العادي يشوف فحص البنك بس
        render_bank_check()