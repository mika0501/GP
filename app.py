import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="GP Cargo 發票資料擷取工具", page_icon="🧾", layout="wide"
)
st.title("🧾 GP Cargo 發票資料自動轉 Excel 工具")
st.caption(
    "已新增 4 個欄位：精準擷取 Remark (PO#)、From、To、Material (SKU#)，並支援累加既有 Excel！"
)


def parse_gp_invoice(file_bytes):
    """解析 GP Cargo 格式發票，提取完整 7 大欄位"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # ================= 1. INVOICE NUMBER =================
    # 相容 6碼-3碼 (如 091726-005) 與 2碼-5碼以上 (如 26-100508)
    invoice_number = ""
    dash_6_match = re.search(r"\b(\d{6}\s*-\s*\d{3})\b", full_text)
    if dash_6_match:
        invoice_number = dash_6_match.group(1).strip()
    else:
        inv_dash_match = re.search(r"\b([0-9]{2}-[0-9]{5,})\b", full_text)
        if inv_dash_match:
            invoice_number = inv_dash_match.group(1).strip()
        else:
            inv_label = re.search(
                r"INVOICE\s*#?\s*[:：]?\s*\n?\s*([A-Za-z0-9\-]+)",
                full_text,
                re.IGNORECASE,
            )
            if inv_label:
                invoice_number = inv_label.group(1).strip()

    # ================= 2. INVOICE DATE =================
    # 優先鎖定打勾處的 Invoice date:，找不到才備援抓取表格 DATE
    date_match = re.search(
        r"Invoice\s*date\s*[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE,
    )
    if not date_match:
        date_match = re.search(
            r"\bDATE\s*\n?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
            full_text,
            re.IGNORECASE,
        )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # ================= 3. TOTAL AMOUNT =================
    # 轉為 Excel 純數值 (float / None)，相容有無小數點
    total_match = re.search(
        r"TOTAL\s*[:：]?\s*\$?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?|\d+)",
        full_text,
        re.IGNORECASE,
    )
    total_amount = None
    if total_match:
        val_str = total_match.group(1).replace(",", "").strip()
        try:
            total_amount = float(val_str)
        except ValueError:
            total_amount = None

    # ================= 4. REMARK =================
    remark = ""
    fee_match = re.search(
        r"Trucking\s+Shipping\s+Fee\s+([A-Za-z0-9]+)\s+from",
        full_text,
        re.IGNORECASE,
    )
    if fee_match:
        remark = fee_match.group(1).strip()
    else:
        po_match = re.search(
            r"PO#?\s*([A-Za-z0-9]+)", full_text, re.IGNORECASE
        )
        if po_match:
            remark = po_match.group(1).strip()

    # ================= 5. FROM =================
    from_val = ""
    from_match = re.search(
        r"\bfrom\s+([A-Za-z0-9\-]+)\s+to\b", full_text, re.IGNORECASE
    )
    if from_match:
        from_val = from_match.group(1).strip()

    # ================= 6. TO =================
    to_val = ""
    to_match = re.search(
        r"\bto\s+([A-Za-z0-9\-]+)(?:\s+|\n|$)", full_text, re.IGNORECASE
    )
    if to_match:
        to_val = to_match.group(1).strip()

    # ================= 7. MATERIAL =================
    material = ""
    sku_match = re.search(
        r"SKU#?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    if sku_match:
        material = sku_match.group(1).strip()

    return {
        "INVOICE NUMBER": invoice_number,
        "INVOICE DATE": invoice_date,
        "TOTAL AMOUNT": total_amount,
        "REMARK": remark,
        "FROM": from_val,
        "TO": to_val,
        "MATERIAL": material,
    }


# ==================== 操作介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳先前整理好的 Excel")
    existing_file = st.file_uploader(
        "若要將新發票資料接續累加至既有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳 PDF 發票檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票（支援多個檔案同時拖入）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data = parse_gp_invoice(file.read())
        records.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(records)

    # 累加合併既有 Excel 邏輯
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            if "INVOICE NUMBER"
