import io
import os
import re
import pandas as pd
import pdfplumber
from PIL import Image, ImageEnhance
import pytesseract
import streamlit as st

# Windows 本機環境 Tesseract 預設路徑備援
if os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )

st.set_page_config(
    page_title="上海騰飛發票表格轉換工具",
    page_icon="📋",
    layout="wide"
)
st.title("📋 上海騰飛發票表格自動轉 Excel 工具（OCR + 逐行精準版）")
st.caption("支援紙本掃描 PDF：先以 300 DPI 影像增強去除雜訊執行 OCR，再以狀態機對齊每個商品明細！")


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


def extract_text_with_ocr(file_bytes):
    """提取文字；若文字過少則強制執行 300 DPI 影像增強 + OCR"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            # 如果頁面純文字少於 60 字元，判定為掃描圖檔，啟動 OCR
            if len(txt.strip()) < 60:
                try:
                    img = page.to_image(resolution=300).original
                    # 影像前處理：轉灰階並增強對比度，去除傳真掃描雜點
                    img = img.convert("L")
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(2.2)
                    ocr_res = pytesseract.image_to_string(img, lang="eng")
                    if ocr_res:
                        txt = ocr_res
                except Exception as e:
                    st.warning(f"OCR 執行提示: {e}")
            full_text += txt + "\n"
    return full_text


def parse_tengfei_invoice(file_bytes):
    # 執行提取（純數位或 OCR 掃描）
    full_text = extract_text_with_ocr(file_bytes)

    # 1. 表頭核心欄位
    inv_num_match = re.search(r"INVOICE\s*NUMBER\s*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    invoice_number = inv_num_match.group(1).strip() if inv_num_match else ""

    inv_date_match = re.search(r"INVOICE\s*DATE\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", full_text, re.IGNORECASE)
    invoice_date = inv_date_match.group(1).strip() if inv_date_match else ""

    # 2. 截取商品表格區間
    table_block_match = re.search(r"HIPPING\s*MARK[\s\S]*?(?=COUNTRY\s*of|SUBTOTAL|THIS\s*TOTAL)", full_text, re.IGNORECASE)
    table_text = table_block_match.group(0) if table_block_match else full_text

    # 預先處理折行的 "30\nPCS"
    table_text = re.sub(r"(\d+)\s*\n\s*PCS", r"\1 PCS", table_text, flags=re.IGNORECASE)

    raw_lines = [line.strip() for line in table_text.splitlines() if line.strip()]

    items = []
    current_item = None

    for line in raw_lines:
        # 排除表頭文字行
        if any(h in line.upper() for h in ["HIPPING MARK", "QUANTITY", "PART NUMBER", "PO NUMBER", "UNIT PRICE", "AMOUNT"]):
            continue

        # 遇到新的數量行 (例如 30 PCS, 120 PCS)
        qty_match = re.search(r"\b(\d+\s*PCS)\b", line, re.IGNORECASE)
        if qty_match and not any(k in line.upper() for k in ["USD", "CASE"]):
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

        # 抓取品名 (帶有 | 且包含英文，但不是 USD 金額行)
        if "|" in line and "USD" not in line.upper():
            desc_clean = line.replace("|", "").strip()
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

    # 校正單價與總額反置問題 (單價通常小於總額)
    for itm in items:
        try:
            up_num = float(itm["UNIT PRICE"].replace("$", ""))
            am_num = float(itm["AMOUNT"].replace("$", ""))
            if up_num > am_num and itm["QUANTITY"]:
                qty_digits = re.search(r"\d+", itm["QUANTITY"])
                if qty_digits and float(qty_digits.group(0)) > 1:
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
        with st.spinner(f"正在執行第 {i+1} 份 PDF 的 OCR 影像掃描辨識中..."):
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
