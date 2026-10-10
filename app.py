import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="上海騰飛發票表格轉換工具", page_icon="📋", layout="wide"
)
st.title("📋 上海騰飛 (Tengfei) 發票表格自動轉 Excel 工具")
st.caption(
    "自動解析發票表頭資訊（發票號碼、日期等）與下方明細清單（數量、料號、PO、品名規格、單價、總額），並匯出完整 Excel！"
)


def parse_tengfei_invoice(file_bytes):
    """解析上海騰飛發票檔案，提取表頭與商品明細"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # ================= 1. 表頭核心欄位 =================
    inv_num_match = re.search(
        r"INVOICE\s*NUMBER\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
    )
    invoice_number = inv_num_match.group(1).strip() if inv_num_match else ""

    inv_date_match = re.search(
        r"INVOICE\s*DATE\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
        full_text,
        re.IGNORECASE,
    )
    invoice_date = inv_date_match.group(1).strip() if inv_date_match else ""

    terms_match = re.search(
        r"PAYMENT\s*TERM\s*([^\n\r]+)", full_text, re.IGNORECASE
    )
    payment_term = terms_match.group(1).strip() if terms_match else ""

    total_amt_match = re.search(
        r"(?:US\$|USD)\s*([0-9,]+\.[0-9]{2})", full_text
    )
    invoice_total = (
        f"${total_amt_match.group(1)}" if total_amt_match else ""
    )

    # ================= 2. 表格項目多列解析 =================
    # 截取表格區域（QUANTITY 到 COUNTRY of / SUBTOTAL 之間）
    table_block_match = re.search(
        r"QUANTITY[\s\S]*?(?=COUNTRY\s*of|SUBTOTAL|THIS\s*TOTAL)",
        full_text,
        re.IGNORECASE,
    )
    table_text = table_block_match.group(0) if table_block_match else full_text

    # 針對每一筆紀錄的模式匹配
    # 項目特徵：數量(數字 PCS) -> 描述說明 -> 兩組 10 碼數字 (Part No & PO No) -> 單價 -> 總金額
    item_pattern = re.compile(
        r"(\d+\s*PCS)[\s\S]*?\|\s*([^\n\r\|]+)[\s\S]*?(\b\d{10}\b)[\s\S]*?(\b\d{10}\b)[\s\S]*?\|\s*USD\s*([0-9,\.]+)[\s\S]*?\|\s*USD\s*([0-9,\.]+)",
        re.IGNORECASE,
    )

    items = []
    for match in item_pattern.finditer(table_text):
        qty = match.group(1).strip()
        description = match.group(2).strip()
        code1 = match.group(3).strip()
        code2 = match.group(4).strip()
        unit_price = f"${match.group(5).strip()}"
        amount = f"${match.group(6).strip()}"

        # 區分 Part No (通常 33 開頭) 與 PO Number (通常 45 開頭)
        if code1.startswith("45") or not code2.startswith("45"):
            po_number = code1 if code1.startswith("45") else code2
            part_number = code2 if code1.startswith("45") else code1
        else:
            part_number = code1
            po_number = code2

        items.append(
            {
                "INVOICE NUMBER": invoice_number,
                "INVOICE DATE": invoice_date,
                "QUANTITY": qty,
                "PART NUMBER": part_number,
                "PO NUMBER": po_number,
                "DESCRIPTION": description,
                "UNIT PRICE": unit_price,
                "AMOUNT": amount,
                "PAYMENT TERM": payment_term,
            }
        )

    return items


# ==================== 網頁介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有 Excel 檔案")
    existing_file = st.file_uploader(
        "若要將新發票明細累加至現有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳 PDF 發票檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳上海騰飛 PDF 發票（支援多個檔案同時上傳）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    all_records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        records = parse_tengfei_invoice(file.read())
        all_records.extend(records)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(all_records)

    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            st.success(
                f"🎉 成功解析 {len(new_df)} 筆明細，並已合併至原有的 Excel！目前累積共 {len(final_df)} 筆資料。"
            )
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(new_df)} 筆商品明細資料！")

    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Invoice_Items")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載轉換後的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="Tengfei_發票明細表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )