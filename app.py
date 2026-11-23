
import io
import re
import unicodedata
import pandas as pd
import pdfplumber
import streamlit as st


# ==================== 網頁設定 ====================

st.set_page_config(
    page_title="GP Cargo 發票資料擷取工具",
    page_icon="🧾",
    layout="wide"
)

st.title("🧾 GP Cargo 發票資料自動轉 Excel 工具")

st.caption(
    "支援批次 PDF 擷取、多個 SKU、含空格的地點名稱、"
    "既有 Excel 累加、重複發票檢查及錯誤偵測。"
)


# ==================== 欄位設定 ====================

FIELDS = [
    "INVOICE NUMBER",
    "INVOICE DATE",
    "TOTAL AMOUNT",
    "REMARK",
    "FROM",
    "TO",
    "MATERIAL"
]

COLUMNS = ["FILENAME"] + FIELDS


# ==================== PDF 解析函式 ====================

def parse_gp_invoice(file_bytes):
    """解析 GP Cargo 發票，提取完整 7 大欄位"""

    full_text = ""

    # 讀取 PDF 文字
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:

            t = page.extract_text(
                x_tolerance=3,
                y_tolerance=3
            )

            if t:
                full_text += t + "\n"

    # 統一特殊符號及全形文字
    full_text = unicodedata.normalize("NFKC", full_text)

    full_text = (
        full_text
        .replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\u00ad", "")
        .replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
        .replace("／", "/")
        .replace("＃", "#")
    )

    # 檢查是否成功擷取文字
    if not full_text.strip():
        raise ValueError("PDF 未擷取到任何文字")

    # 將換行及多餘空格統一
    flat_text = re.sub(
        r"\s+", " ", full_text
    ).strip()


    # ================= 1. INVOICE NUMBER =================

    invoice_number = ""

    inv_match = re.search(
        r"INVOICE\s*#?\s*DATE\s*"
        r"(\d{6}\s*-\s*\d{3}|\d{2}\s*-\s*\d{5,})",
        flat_text,
        re.IGNORECASE
    )

    # 備援搜尋
    if not inv_match:
        inv_match = re.search(
            r"\b(\d{6}\s*-\s*\d{3}|\d{2}\s*-\s*\d{5,})\b",
            flat_text
        )

    if inv_match:
        invoice_number = re.sub(
            r"\s+", "", inv_match.group(1)
        )


    # ================= 2. INVOICE DATE =================

    date_match = re.search(
        r"Invoice\s*date\s*[:：]?\s*"
        r"(\d{1,2}/\d{1,2}/\d{4})",
        flat_text,
        re.IGNORECASE
    )

    # 備援：從發票號碼後方尋找日期
    if not date_match:
        date_match = re.search(
            r"\b\d{2,6}\s*-\s*\d{3,}\s+"
            r"(\d{1,2}/\d{1,2}/\d{4})",
            flat_text
        )

    invoice_date = (
        date_match.group(1).strip()
        if date_match else ""
    )


    # ================= 3. TOTAL AMOUNT =================

    total_amount = None

    total_match = re.search(
        r"\bTOTAL\s*[:：]?\s*\$?\s*"
        r"(\d[\d,]*(?:\.\d{1,2})?)",
        flat_text,
        re.IGNORECASE
    )

    if total_match:
        try:
            total_amount = float(
                total_match.group(1).replace(",", "")
            )
        except ValueError:
            total_amount = None


    # ================= 4. REMARK =================

    remark = ""

    # 優先從 PO# 擷取
    po_match = re.search(
        r"\bPO\s*#\s*[:：]?\s*([A-Za-z0-9\-]+)",
        flat_text,
        re.IGNORECASE
    )

    if po_match:
        remark = po_match.group(1).strip()


    # ================= 5 & 6. FROM / TO =================

    from_val = ""
    to_val = ""

    # 同時擷取貨運編號、出發地、目的地
    route_match = re.search(
        r"Trucking\s+Shipping\s+Fee\s+"
        r"([A-Za-z0-9\-]+)\s+"
        r"from\s+(.+?)\s+"
        r"to\s+(.+?)"
        r"(?=\s*\$\s*\d"
        r"|\s+\d[\d,]*\.\d{2}"
        r"|\s+(?:\d+\s+)?Processing\s+fee"
        r"|$)",
        flat_text,
        re.IGNORECASE
    )

    if route_match:

        if not remark:
            remark = route_match.group(1).strip()

        from_val = route_match.group(2).strip()
        to_val = route_match.group(3).strip()

    # REMARK 備援擷取
    if not remark:

        fee_match = re.search(
            r"Trucking\s+Shipping\s+Fee\s+"
            r"([A-Za-z0-9\-]+)",
            flat_text,
            re.IGNORECASE
        )

        if fee_match:
            remark = fee_match.group(1).strip()


    # ================= 7. MATERIAL =================

    material = ""

    # 支援單一及多個 SKU
    sku_match = re.search(
        r"\bSKU\s*#?\s*[:：]?\s*"
        r"([A-Za-z0-9\-]+"
        r"(?:\s*/\s*[A-Za-z0-9\-]+)*)",
        flat_text,
        re.IGNORECASE
    )

    if sku_match:

        material = re.sub(
            r"\s*/\s*",
            "/",
            sku_match.group(1)
        ).strip()


    # ================= 回傳結果 =================

    data = {
        "INVOICE NUMBER": invoice_number,
        "INVOICE DATE": invoice_date,
        "TOTAL AMOUNT": total_amount,
        "REMARK": remark,
        "FROM": from_val,
        "TO": to_val,
        "MATERIAL": material,
    }

    return data, full_text


# ==================== 資料檢查函式 ====================

def is_empty(value):
    """判斷欄位是否為空白"""

    return pd.isna(value) or str(value).strip() == ""


def normalize_invoice_number(value):
    """統一發票號碼格式，避免重複資料"""

    if is_empty(value):
        return ""

    value = unicodedata.normalize("NFKC", str(value))

    value = (
        value
        .replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
    )

    return re.sub(r"\s+", "", value).upper()


# ==================== Excel 合併函式 ====================

def merge_invoice_data(old_df, new_df):
    """合併新舊發票，避免空白資料覆蓋原有資料"""

    combined_df = pd.concat(
        [old_df, new_df],
        ignore_index=True
    )

    # 確保必要欄位存在
    for column in COLUMNS:
        if column not in combined_df.columns:
            combined_df[column] = None

    combined_df["INVOICE NUMBER"] = (
        combined_df["INVOICE NUMBER"]
        .apply(normalize_invoice_number)
    )

    merged_records = []
    invoice_positions = {}

    for _, row in combined_df.iterrows():

        record = row.to_dict()
        invoice_number = record["INVOICE NUMBER"]

        # 無發票號碼的資料保留，不進行去重
        if not invoice_number:
            merged_records.append(record)
            continue

        # 新發票
        if invoice_number not in invoice_positions:

            invoice_positions[invoice_number] = len(
                merged_records
            )

            merged_records.append(record)

        # 重複發票：只更新非空白欄位
        else:

            position = invoice_positions[invoice_number]

            for key, value in record.items():

                if not is_empty(value):
                    merged_records[position][key] = value

    return pd.DataFrame(
        merged_records,
        columns=combined_df.columns
    )


# ==================== 操作介面 ====================

col1, col2 = st.columns(2)

with col1:

    st.markdown(
        "### 步驟 1：(選填) 上傳先前整理好的 Excel"
    )

    existing_file = st.file_uploader(
        "若要將新發票資料接續累加至既有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel"
    )


with col2:

    st.markdown("### 步驟 2：上傳 PDF 發票檔案")

    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票（支援多個檔案同時拖入）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files"
    )


# ==================== Debug 設定 ====================

show_debug = st.checkbox(
    "🔍 顯示 PDF 原始擷取文字（Debug 模式）",
    value=False
)


# ==================== 開始解析 ====================

if uploaded_files:

    records = []

    failed_files = []

    incomplete_files = []

    complete_count = 0

    progress_bar = st.progress(0)


    # ==================== 批次讀取 ====================

    for i, file in enumerate(uploaded_files):

        try:

            data, raw_text = parse_gp_invoice(
                file.getvalue()
            )

            # 顯示原始 PDF 文字
            if show_debug:

                with st.expander(
                    f"🔍 {file.name} - 原始擷取文字"
                ):

                    st.text(raw_text)

            # 檢查是否有欄位未成功擷取
            missing = [
                key
                for key, value in data.items()
                if is_empty(value)
            ]

            if missing:

                incomplete_files.append(file.name)

                st.warning(
                    f"⚠️ {file.name}：以下欄位未成功擷取："
                    + ", ".join(missing)
                )

            else:

                complete_count += 1

            # 加入檔案名稱
            data_with_name = {
                "FILENAME": file.name,
                **data
            }

            records.append(data_with_name)

        except Exception as e:

            failed_files.append(file.name)

            st.error(
                f"❌ {file.name} 讀取失敗：{e}"
            )

        finally:

            progress_bar.progress(
                (i + 1) / len(uploaded_files)
            )


    # 若所有 PDF 均失敗
    if not records:

        st.error(
            "沒有成功讀取任何發票，請檢查 PDF 檔案格式。"
        )

        st.stop()


    # ==================== 建立新資料表 ====================

    new_df = pd.DataFrame(
        records,
        columns=COLUMNS
    )


    # ==================== 合併既有 Excel ====================

    if existing_file:

        try:

            old_df = pd.read_excel(
                io.BytesIO(existing_file.getvalue()),
                dtype=str
            )

            # 確認既有 Excel 包含發票號碼
            if "INVOICE NUMBER" not in old_df.columns:

                raise ValueError(
                    "既有 Excel 缺少 INVOICE NUMBER 欄位"
                )

            final_df = merge_invoice_data(
                old_df,
                new_df
            )

            st.success(
                f"已合併既有 Excel，"
                f"目前共 {len(final_df)} 筆資料。"
            )

        except Exception as e:

            st.error(
                f"讀取既有 Excel 失敗：{e}"
            )

            st.warning(
                "以下結果僅包含本次上傳的 PDF 資料，"
                "尚未合併舊 Excel。"
            )

            final_df = merge_invoice_data(
                pd.DataFrame(columns=COLUMNS),
                new_df
            )

    else:

        final_df = merge_invoice_data(
            pd.DataFrame(columns=COLUMNS),
            new_df
        )


    # ==================== 整理欄位格式 ====================

    # 將金額轉為 Excel 數值
    final_df["TOTAL AMOUNT"] = pd.to_numeric(
        final_df["TOTAL AMOUNT"]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False),
        errors="coerce"
    )

    # 固定主要欄位順序
    extra_columns = [
        col for col in final_df.columns
        if col not in COLUMNS
    ]

    final_df = final_df[
        COLUMNS + extra_columns
    ]


    # ==================== 顯示處理結果 ====================

    st.subheader("📋 發票處理結果")

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "上傳 PDF",
        len(uploaded_files)
    )

    m2.metric(
        "完整擷取",
        complete_count
    )

    m3.metric(
        "部分欄位缺失",
        len(incomplete_files)
    )

    m4.metric(
        "讀取失敗",
        len(failed_files)
    )


    # ==================== 預覽及手動修改 ====================

    st.subheader("📊 擷取結果預覽")

    st.caption(
        "可以直接點擊表格中的儲存格修改資料，"
        "修改後的內容會一起匯出至 Excel。"
    )

    edited_df = st.data_editor(
        final_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "TOTAL AMOUNT": st.column_config.NumberColumn(
                "TOTAL AMOUNT",
                format="$%.2f"
            )
        },
        key="invoice_editor"
    )


    # ==================== 金額統計 ====================

    total_sum = edited_df["TOTAL AMOUNT"].sum()

    st.metric(
        "發票總金額",
        f"${total_sum:,.2f}"
    )


    # ==================== 匯出 Excel ====================

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        edited_df.to_excel(
            writer,
            index=False,
            sheet_name="GP_Invoices"
        )

        worksheet = writer.sheets["GP_Invoices"]

        # 固定第一列
        worksheet.freeze_panes = "A2"

        # 加入篩選功能
        worksheet.auto_filter.ref = worksheet.dimensions

        # 自動調整欄位寬度
        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = column_cells[0].column_letter

            for cell in column_cells:

                if cell.value is not None:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            worksheet.column_dimensions[
                column_letter
            ].width = min(max_length + 4, 80)

        # 設定金額欄位格式
        amount_column = (
            edited_df.columns.get_loc("TOTAL AMOUNT") + 1
        )

        for row_number in range(
            2, worksheet.max_row + 1
        ):

            worksheet.cell(
                row=row_number,
                column=amount_column
            ).number_format = '"$"#,##0.00'


    excel_data = output.getvalue()


    # ==================== 下載按鈕 ====================

    st.download_button(
        label="📥 下載整理好的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="GP_Cargo_發票整理.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        type="primary"
    )

else:

    st.info(
        "請先上傳 PDF 發票，即可開始自動擷取資料。"
    )
