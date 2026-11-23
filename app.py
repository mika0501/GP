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
    "支援 GP CARGO 含空格地點、多組斜線 SKU#、純數值金額，並支援累加既有 Excel！"
)


def extract_clean_text(file_bytes):
    """嘗試多種容差參數提取 PDF 文字，並清理隱形控制字元"""
    text_standard = ""
    text_tolerant = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t1 = page.extract_text() or ""
            text_standard += t1 + "\n"
            t2 = page.extract_text(x_tolerance=5, y_tolerance=5) or ""
            text_tolerant += t2 + "\n"

    # 移除空字元 \x00 與異常控制字元
    text_standard = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text_standard)
    text_tolerant = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text_tolerant)

    return text_standard, text_tolerant


def parse_text_fields(full_text):
    """從文字中提取 7 大欄位"""
    # 1. INVOICE NUMBER
    invoice_number = ""
    dash_6_match = re.search(r"\b(\d{6}\s*-\s*\d{3})\b", full_text)
    if dash_6_match:
        invoice_number = dash_6_match.group(1).replace(" ", "").strip()
    else:
        inv_dash_match = re.search(r"\b([0-9]{2}\s*-\s*[0-9]{5,})\b", full_text)
        if inv_dash_match:
            invoice_number = inv_dash_match.group(1).replace(" ", "").strip()
        else:
            inv_label = re.search(
                r"INVOICE\s*#?\s*[:：]?\s*\n?\s*([A-Za-z0-9\-]+)",
                full_text,
                re.IGNORECASE,
            )
            if inv_label:
                invoice_number = inv_label.group(1).strip()

    # 2. INVOICE DATE (優先鎖定打勾處的 Invoice date:)
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

    # 3. TOTAL AMOUNT
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

    # 4. REMARK
    remark = ""
    fee_match = re.search(
        r"Trucking\s+Shipping\s+Fee\s+([A-Za-z0-9\-]+)\s+from",
        full_text,
        re.IGNORECASE,
    )
    if fee_match:
        remark = fee_match.group(1).strip()
    else:
        po_match = re.search(
            r"PO#?\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
        )
        if po_match:
            remark = po_match.group(1).strip()

    # 5 & 6. FROM & TO
    from_val = ""
    to_val = ""
    route_match = re.search(
        r"\bfrom\s+([A-Za-z0-9\-\s]+?)\s+to\s+([A-Za-z0-9\-\s]+?)(?=\s+\$|\n|$)",
        full_text,
        re.IGNORECASE,
    )
    if route_match:
        from_val = re.sub(r"\s+", " ", route_match.group(1)).strip()
        to_val = re.sub(r"\s+", " ", route_match.group(2)).strip()

    # 7. MATERIAL
    material = ""
    sku_match = re.search(
        r"SKU#?\s*([A-Za-z0-9\-/]+)", full_text, re.IGNORECASE
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


def parse_gp_invoice(file_bytes):
    """解析 GP Cargo 格式發票，並回傳解析結果與原始文字供除錯"""
    text_std, text_tol = extract_clean_text(file_bytes)

    # 先嘗試標準文字解析
    res = parse_text_fields(text_std)

    # 若標準模式沒抓到發票號碼或日期，改用高容差文字重試
    if not res["INVOICE NUMBER"] and not res["INVOICE DATE"]:
        res_tol = parse_text_fields(text_tol)
        if res_tol["INVOICE NUMBER"] or res_tol["INVOICE DATE"]:
            res = res_tol

    # 判斷完全無有效文字的情況
    if not text_std.strip() and not text_tol.strip():
        res["REMARK"] = "⚠️ 無法讀取文字（字型編碼加密或掃描圖檔）"

    raw_preview = text_std.strip() if text_std.strip() else text_tol.strip()
    return res, raw_preview


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
    debug_texts = {}
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data, raw_txt = parse_gp_invoice(file.read())
        data_with_name = {"FILENAME": file.name, **data}
        records.append(data_with_name)
        debug_texts[file.name] = raw_txt
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(records)

    # 累加合併既有 Excel 邏輯
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            if "INVOICE NUMBER" in final_df.columns:
                has_inv = final_df["INVOICE NUMBER"].notna() & (
                    final_df["INVOICE NUMBER"].astype(str).str.strip() != ""
                )
                df_valid = final_df[has_inv].drop_duplicates(
                    subset=["INVOICE NUMBER"], keep="last"
                )
                df_empty = final_df[~has_inv]
                final_df = pd.concat([df_valid, df_empty], ignore_index=True)
            st.success(
                f"🎉 成功解析 {len(records)} 筆發票，並累加至現有 Excel！目前共 {len(final_df)} 筆資料。"
            )
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(records)} 筆發票資料！")

    # 預覽表格
    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    # 匯出 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="GP_Invoices")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="GP_Cargo_發票整理.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    # 除錯透視鏡：針對沒抓到的檔案顯示底層讀出的文字
    empty_files = [
        r["FILENAME"]
        for r in records
        if not r["INVOICE NUMBER"] and not r["INVOICE DATE"]
    ]
    if empty_files:
        with st.expander(f"🔍 異常檔案透視（共 {len(empty_files)} 份未抓到資料，點此查看底層讀出的文字）"):
            selected_err = st.selectbox("選擇要檢查的檔案：", empty_files)
            st.text_area(
                "PDF 底層實際讀出的文字內容：",
                value=debug_texts.get(selected_err, "") or "（完全空白，無任何可讀取字元）",
                height=250,
            )
