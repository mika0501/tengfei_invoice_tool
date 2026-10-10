import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="上海騰飛發票表格轉換工具",
    page_icon="📋",
    layout="wide"
)
st.title("📋 上海騰飛發票表格自動轉 Excel 工具")
st.caption("已修復：採用逐行狀態機精準對齊，徹底解決品名空白、數量錯位與金額反置問題！")

def clean_price(val_str):
    """將金額統一轉為乾淨的小數點格式"""
    if not val_str:
        return ""
    val = val_str.replace(" ", "").strip()
    if "," in val and "." in val:
        val = val.replace(",", "")
    else:
        val = val.replace(",", ".")
    return val

def parse_tengfei_invoice(file_bytes):
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # 1. 表頭核心欄位
    inv_num_match = re.search(r"INVOICE\s*NUMBER\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    invoice_number = inv_num_match.group(1).strip() if inv_num_match else ""

    inv_date_match = re.search(r"INVOICE\s*DATE\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", full_text, re.IGNORECASE)
    invoice_date = inv_date_match.group(1).strip() if inv_date_match else ""

    # 2. 截取商品表格區間
    table_block_match = re.search(r"HIPPING\s*MARK[\s\S]*?(?=COUNTRY\s*of|SUBTOTAL|THIS\s*TOTAL)", full_text, re.IGNORECASE)
    table_text = table_block_match.group(0) if table_block_match else full_text

    # 預先處理換行的 "30\nPCS"
    table_text = re.sub(r"(\d+)\s*\n\s*PCS", r"\1 PCS", table_text, flags=re.IGNORECASE)

    raw_lines = [line.strip() for line in table_text.splitlines() if line.strip()]

    items = []
    current_item = None

    for line in raw_lines:
        # 排除表頭與雜訊行
        if any(h in line.upper() for h in ["HIPPING MARK", "QUANTITY", "PART NUMBER", "PO NUMBER", "UNIT PRICE", "AMOUNT"]):
            continue

        # 遇到新的數量行 (例如 30 PCS, 120 PCS)
        qty_match = re.match(r"^(\d+\s*PCS)$", line, re.IGNORECASE)
        if qty_match:
            # 結算上一筆
            if current_item and current_item["QUANTITY"]:
                items.append(current_item)
            current_item = {
                "INVOICE NUMBER": invoice_number,
                "INVOICE DATE": invoice_date,
                "QUANTITY": qty_match.group(1).upper(),
                "PART NUMBER": "",
                "PO NUMBER": "",
                "DESCRIPTION": "",
                "UNIT PRICE": "",
                "AMOUNT": "",
            }
            continue

        if not current_item:
            continue

        # 抓取品名 (帶有 | 且包含英文字母，但不是 USD 金額行)
        if "|" in line and "USD" not in line.upper():
            desc_clean = line.replace("|", "").strip()
            # 濾除雜訊
            if desc_clean and not desc_clean.startswith("45") and not desc_clean.startswith("33"):
                if not current_item["DESCRIPTION"]:
                    current_item["DESCRIPTION"] = desc_clean
                else:
                    current_item["DESCRIPTION"] += " " + desc_clean
            continue

        # 抓取 10 碼代碼 (45 開頭為 PO，33 開頭為 Part No)
        codes = re.findall(r"\b\d{10}\b", line)
        if codes:
            for c in codes:
                if c.startswith("45"):
                    current_item["PO NUMBER"] = c
                elif c.startswith("33"):
                    current_item["PART NUMBER"] = c
                else:
                    if not current_item["PART NUMBER"]:
                        current_item["PART NUMBER"] = c
                    elif not current_item["PO NUMBER"]:
                        current_item["PO NUMBER"] = c
            continue

        # 抓取 USD 金額
        if "USD" in line.upper():
            prices = re.findall(r"USD\s*([0-9,\.]+)", line, re.IGNORECASE)
            for p in prices:
                clean_p = f"${clean_price(p)}"
                if not current_item["UNIT PRICE"]:
                    current_item["UNIT PRICE"] = clean_p
                elif not current_item["AMOUNT"]:
                    current_item["AMOUNT"] = clean_p
            continue

    # 結算最後一筆
    if current_item and current_item["QUANTITY"]:
        items.append(current_item)

    # 後處理：校正單價與總額顛倒問題 (單價通常小於等於總金額)
    for itm in items:
        try:
            up_num = float(itm["UNIT PRICE"].replace("$", ""))
            am_num = float(itm["AMOUNT"].replace("$", ""))
            if up_num > am_num and itm["QUANTITY"]:
                qty_num = float(re.search(r"\d+", itm["QUANTITY"]).group(0))
                if qty_num > 1:
                    # 顛倒了，對調
                    itm["UNIT PRICE"], itm["AMOUNT"] = itm["AMOUNT"], itm["UNIT PRICE"]
        except Exception:
            pass

    return items

# ==================== 網頁介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有 Excel 檔案")
    existing_file = st.file_uploader(
        "若要將新發票明細累加至現有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel"
    )

with col2:
    st.markdown("### 步驟 2：上傳 PDF 發票檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳上海騰飛 PDF 發票（支援多個檔案同時上傳）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files"
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
            st.success(f"🎉 成功解析 {len(new_df)} 筆明細，並已合併至原有的 Excel！目前累積共 {len(final_df)} 筆資料。")
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
        type="primary"
    )
