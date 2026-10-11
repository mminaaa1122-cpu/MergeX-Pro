import streamlit as st
import pandas as pd
from io import BytesIO


BANK_REQUIRED_COLUMNS = [
    "Serial",
    "SSNumber",
    "Customer Name",
    "Mobile",
    "HOME NUM",
    "Address",
]

BATCH_SEARCH_COLUMN = "SSNumber"


def normalize_ssnumber(series):
    """
    تنظيف SSNumber للمقارنة فقط.
    لا يغير قيمة الملف الأصلي.
    """
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def read_ssnumbers_from_excel(uploaded_file):
    """
    قراءة SSNumber من جميع الشيتات الموجودة داخل Batch.
    ترجع Set سريع للبحث.
    """
    ssnumbers = set()

    try:
        uploaded_file.seek(0)

        with pd.ExcelFile(uploaded_file) as excel_file:

            for sheet_name in excel_file.sheet_names:

                df = pd.read_excel(
                    excel_file,
                    sheet_name=sheet_name,
                    dtype=str
                )

                df.columns = (
                    df.columns
                    .astype(str)
                    .str.strip()
                )

                if BATCH_SEARCH_COLUMN not in df.columns:
                    continue

                values = normalize_ssnumber(
                    df[BATCH_SEARCH_COLUMN]
                )

                values = values[values != ""]

                ssnumbers.update(values.tolist())

        return ssnumbers

    except Exception as e:
        raise Exception(
            f"تعذر قراءة الملف {uploaded_file.name}: {e}"
        )


def create_excel_download(df, sheet_name="Clean_Batch"):
    """
    إنشاء Excel داخل الذاكرة للتحميل.
    """
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name
        )

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

        for col_num, column_name in enumerate(df.columns):

            worksheet.write(
                0,
                col_num,
                column_name,
                header_format
            )

        for col_num, column_name in enumerate(df.columns):

            max_length = len(str(column_name))

            if not df.empty:

                sample_values = (
                    df[column_name]
                    .astype(str)
                    .head(1000)
                )

                for value in sample_values:

                    max_length = max(
                        max_length,
                        len(value)
                    )

            width = min(
                max(max_length + 2, 12),
                40
            )

            worksheet.set_column(
                col_num,
                col_num,
                width
            )

        worksheet.freeze_panes(1, 0)

        worksheet.autofilter(
            0,
            0,
            max(len(df), 1),
            len(df.columns) - 1
        )

    output.seek(0)

    return output.getvalue()


def create_duplicate_report(df):
    """
    إنشاء ملف Duplicate Report.
    """
    return create_excel_download(
        df,
        sheet_name="Duplicates"
    )


def validate_bank_file(df):
    """
    التأكد أن ملف البنك يحتوي على الأعمدة المطلوبة.
    """
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    missing_columns = [
        col
        for col in BANK_REQUIRED_COLUMNS
        if col not in df.columns
    ]

    return missing_columns


def process_batch_cleaner(
    bank_file,
    batch_files_by_year
):
    """
    تنفيذ عملية البحث كاملة.

    يبحث عن SSNumber الخاص بكل سجل
    في جميع الـ Batches المختارة.

    الملفات الأصلية لا يتم تعديلها.
    """

    # ==================================================
    # 1. قراءة ملف البنك
    # ==================================================

    bank_file.seek(0)

    bank_df = pd.read_excel(
        bank_file,
        dtype=str
    )

    bank_df.columns = (
        bank_df.columns
        .astype(str)
        .str.strip()
    )

    missing_columns = validate_bank_file(
        bank_df
    )

    if missing_columns:

        raise ValueError(
            "ملف البنك ناقص الأعمدة التالية: "
            + ", ".join(missing_columns)
        )

    # الاحتفاظ بنفس أعمدة البنك وترتيبها
    original_columns = bank_df.columns.tolist()

    # SSNumber للمقارنة فقط
    bank_search_key = normalize_ssnumber(
        bank_df[BATCH_SEARCH_COLUMN]
    )

    # ==================================================
    # 2. قراءة كل Batch
    # ==================================================

    all_existing_ssnumbers = set()

    year_ssnumbers = {}

    for year, batch_file in batch_files_by_year.items():

        if batch_file is None:
            continue

        try:

            ssnumbers = read_ssnumbers_from_excel(
                batch_file
            )

            year_ssnumbers[year] = ssnumbers

            all_existing_ssnumbers.update(
                ssnumbers
            )

        except Exception as e:

            raise Exception(
                f"خطأ أثناء قراءة Batch {year}: {e}"
            )

    # ==================================================
    # 3. تحديد الـ Duplicates
    # ==================================================

    duplicate_mask = bank_search_key.isin(
        all_existing_ssnumbers
    )

    duplicate_indices = bank_df.index[
        duplicate_mask
    ]

    clean_indices = bank_df.index[
        ~duplicate_mask
    ]

    # ==================================================
    # 4. Clean Batch
    # ==================================================

    clean_df = bank_df.loc[
        clean_indices,
        original_columns
    ].copy()

    clean_df.reset_index(
        drop=True,
        inplace=True
    )

    # ==================================================
    # 5. Duplicate Report
    # ==================================================

    duplicate_df = bank_df.loc[
        duplicate_indices,
        original_columns
    ].copy()

    duplicate_df.reset_index(
        drop=True,
        inplace=True
    )

    if not duplicate_df.empty:

        duplicate_search_keys = normalize_ssnumber(
            duplicate_df[BATCH_SEARCH_COLUMN]
        )

        found_in_values = []

        for key in duplicate_search_keys:

            found_years = []

            for year, year_set in year_ssnumbers.items():

                if key and key in year_set:
                    found_years.append(year)

            found_in_values.append(
                ", ".join(found_years)
            )

        duplicate_df["Found In"] = (
            found_in_values
        )

    else:

        duplicate_df["Found In"] = pd.Series(
            dtype=str
        )

    # ==================================================
    # 6. Summary
    # ==================================================

    summary = {
        "bank_total": len(bank_df),
        "duplicates": len(duplicate_df),
        "clean_records": len(clean_df),
        "batches_checked": list(
            batch_files_by_year.keys()
        ),
    }

    return (
        clean_df,
        duplicate_df,
        summary
    )


def render_batch_cleaner():
    """
    واجهة Batch Cleaner كاملة.
    """

    st.markdown(
        "<h1>🧹 Batch Cleaner</h1>",
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style="
            text-align:center;
            color:#94a3b8;
            font-size:1.05rem;
        ">
        تنظيف Batch البنك ومقارنة الـ SSNumber
        مع الـ Batches السابقة بدون تعديل الملفات الأصلية
        </p>
        """,
        unsafe_allow_html=True
    )

    st.markdown("---")

    # ==================================================
    # Bank File
    # ==================================================

    st.subheader("📥 Bank File")

    bank_file = st.file_uploader(
        "ارفع ملف البنك",
        type=["xlsx", "xls"],
        key="batch_cleaner_bank_file",
        help=(
            "ملف البنك يجب أن يحتوي على "
            "Serial و SSNumber وباقي بيانات العميل"
        )
    )

    if bank_file is not None:

        try:

            bank_file.seek(0)

            preview_df = pd.read_excel(
                bank_file,
                dtype=str,
                nrows=5
            )

            preview_df.columns = (
                preview_df.columns
                .astype(str)
                .str.strip()
            )

            missing = validate_bank_file(
                preview_df
            )

            if missing:

                st.error(
                    "❌ ملف البنك ناقص الأعمدة: "
                    + ", ".join(missing)
                )

            else:

                detected_columns = " | ".join(
                    preview_df.columns.tolist()
                )

                st.success(
                    "✅ تم التعرف على ملف البنك بنجاح"
                )

                st.markdown(
                    f"""
                    <div style="
                        background:rgba(34,197,94,0.08);
                        border:1px solid rgba(34,197,94,0.3);
                        border-radius:12px;
                        padding:12px 16px;
                        margin:10px 0 20px 0;
                    ">

                        <strong>Detected columns:</strong><br>

                        {detected_columns}

                    </div>
                    """,
                    unsafe_allow_html=True
                )

                st.caption(
                    "🔑 Search Key: SSNumber"
                )

        except Exception as e:

            st.error(
                f"❌ تعذر قراءة ملف البنك: {e}"
            )

    st.markdown("---")

    # ==================================================
    # Existing Batches
    # ==================================================

    st.subheader("📚 Existing Batches")

    st.caption(
        "اختار الـ Batches التي تريد البحث فيها. "
        "الملفات الأصلية لن يتم تعديلها."
    )

    selected_years = {}

    years = [
        "2023",
        "2024",
        "2025",
        "2026"
    ]

    for year in years:

        col1, col2 = st.columns([1, 4])

        with col1:

            enabled = st.checkbox(
                year,
                value=True,
                key=f"batch_year_{year}"
            )

        with col2:

            if enabled:

                uploaded_batch = st.file_uploader(
                    f"📂 Batch {year}",
                    type=["xlsx", "xls"],
                    key=f"batch_file_{year}",
                    help=(
                        f"ارفع ملف أو ملف Batch سنة {year}"
                    )
                )

                if uploaded_batch is not None:

                    selected_years[year] = (
                        uploaded_batch
                    )

                    st.success(
                        f"✅ Batch {year} جاهز للبحث"
                    )

    st.markdown("---")

    # ==================================================
    # Check Button
    # ==================================================

    check_button = st.button(
        "🔍 Check Duplicates",
        use_container_width=True,
        type="primary"
    )

    if check_button:

        if bank_file is None:

            st.error(
                "❌ ارفع ملف البنك أولًا."
            )

            return

        if not selected_years:

            st.error(
                "❌ اختار وارفع Batch واحد على الأقل."
            )

            return

        # التأكد أن كل Batch مختار تم رفعه

        missing_selected_files = []

        for year in years:

            checkbox_key = (
                f"batch_year_{year}"
            )

            if st.session_state.get(
                checkbox_key,
                False
            ):

                if year not in selected_years:

                    missing_selected_files.append(
                        year
                    )

        if missing_selected_files:

            st.warning(
                "⚠️ اخترت السنوات التالية ولكن لم ترفع ملفاتها: "
                + ", ".join(
                    missing_selected_files
                )
            )

            return

        # ==================================================
        # Processing
        # ==================================================

        progress = st.progress(0)

        status = st.empty()

        try:

            status.info(
                "⏳ جاري قراءة الـ Batches..."
            )

            progress.progress(20)

            (
                clean_df,
                duplicate_df,
                summary
            ) = process_batch_cleaner(
                bank_file,
                selected_years
            )

            progress.progress(80)

            status.info(
                "⏳ جاري تجهيز النتائج..."
            )

            st.session_state[
                "batch_clean_result"
            ] = clean_df

            st.session_state[
                "batch_duplicate_result"
            ] = duplicate_df

            st.session_state[
                "batch_clean_summary"
            ] = summary

            progress.progress(100)

            status.success(
                "✅ تم فحص الـ Batch بنجاح"
            )

        except Exception as e:

            progress.empty()

            st.error(
                f"❌ حدث خطأ أثناء الفحص:\n\n{e}"
            )

            return

    # ==================================================
    # Results
    # ==================================================

    if (
        "batch_clean_result"
        not in st.session_state
    ):
        return

    clean_df = st.session_state[
        "batch_clean_result"
    ]

    duplicate_df = st.session_state[
        "batch_duplicate_result"
    ]

    summary = st.session_state[
        "batch_clean_summary"
    ]

    st.markdown("---")

    st.subheader("📊 Results")

    # ==================================================
    # Metrics
    # ==================================================

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "إجمالي البنك",
            f"{summary['bank_total']:,}"
        )

    with col2:

        st.metric(
            "Duplicates",
            f"{summary['duplicates']:,}"
        )

    with col3:

        st.metric(
            "Clean Records",
            f"{summary['clean_records']:,}"
        )

    st.markdown("---")

    # ==================================================
    # Duplicate Result
    # ==================================================

    st.subheader("🟥 Duplicate Records")

    if duplicate_df.empty:

        st.success(
            "🎉 لا يوجد أي Duplicate في الـ Batches المحددة."
        )

    else:

        st.warning(
            f"تم العثور على "
            f"**{len(duplicate_df):,}** "
            f"سجل موجود بالفعل."
        )

        with st.expander(
            "👁️ معاينة الـ Duplicates",
            expanded=False
        ):

            st.dataframe(
                duplicate_df.head(100),
                use_container_width=True,
                height=400
            )

        duplicate_excel = (
            create_duplicate_report(
                duplicate_df
            )
        )

        st.download_button(
            label="📥 تحميل Duplicate Records",
            data=duplicate_excel,
            file_name="Duplicates.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            key="download_duplicates"
        )

    st.markdown("---")

    # ==================================================
    # Clean Batch
    # ==================================================

    st.subheader("🟢 Clean Batch")

    st.success(
        f"الملف النظيف يحتوي على "
        f"**{len(clean_df):,}** سجل."
    )

    with st.expander(
        "👁️ معاينة Clean Batch",
        expanded=False
    ):

        st.dataframe(
            clean_df.head(100),
            use_container_width=True,
            height=400
        )

    clean_excel = create_excel_download(
        clean_df,
        sheet_name="Clean_Batch"
    )

    st.download_button(
        label="📥 تحميل Clean Batch",
        data=clean_excel,
        file_name="Clean_Batch.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
        key="download_clean_batch"
    )

    st.markdown("---")

    st.info(
        "🔒 الملفات الأصلية للـ Batches لم يتم تعديلها. "
        "تم استخدام SSNumber للمقارنة فقط."
    )