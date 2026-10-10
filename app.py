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
    layout="wide",
)
st.title("📋 上海騰飛發票表格自動轉 Excel 工具（OCR 掃描 + 逐行精準對齊版）")
st.caption(
    "雙核心驅動：300 DPI 影像前處理 OCR 掃描 ＋ 逐行狀態機演算法，徹底搞定紙本掃描錯位與漏抓！"
)


# ==================== 1. 資料清洗小工具 ====================
def clean_price(val_str):
    """將金額字串統一清理為標準小數點格式"""
    if not val_str:
        return ""
    val = val_str.replace(" ", "").strip()
    if "," in val and "." in val:
        val = val.replace(",", "")
    else:
        val = val.replace(",", ".")
    return val


# ==================== 2. OCR 影像辨識層 ====================
def extract_text_with_ocr(file_bytes):
    """
    雙模文字提取：
    數位 PDF 直接讀取；紙本掃描 PDF 強制執行 300 DPI 影像增強 + OCR 辨識
    """
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            # 若頁面文字少於 100 字元，判定為掃描檔或傳真件，啟動 OCR
            if len(txt.strip()) < 100:
                try:
                    # 轉為 300 DPI 高解析度圖
                    img = page.to_image(resolution=300).original
                    # 灰階化 + 對比度增強 (2.2倍)，消除紙張底噪
                    img = img.convert("L")
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(2.2)
                    ocr_res = pytesseract.image_to_string(img, lang="eng")
                    if ocr_res:
                        txt = ocr_res
                except Exception as e:
                    st.warning(f"OCR 影像掃描提示: {e}")
            full_text += txt + "\n"
    return full_text


# ==================== 3. 逐行狀態機精準對齊層 ====================
def parse_tengfei_invoice(file_bytes):
    """解析發票：先進行 OCR 掃描提取，再透過狀態機對齊所有商品品項"""
    full_text = extract_text_with_ocr(file_bytes)

    # 1. 抓取表頭核心欄位
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

    # 2. 定位商品表格核心區域
    table_block_match = re.search(
        r"HIPPING\s*MARK[\s\S]*?(?=COUNTRY\s*of|SUBTOTAL|THIS\s*TOTAL)",
        full_text,
        re.IGNORECASE,
    )
    table_text = table_block_match.group(0) if table_block_match else full_text

    # 預先處理折行的數量單位（例如 30\nPCS 轉為 30 PCS）
    table_text = re.sub(
        r"(\d+)\s*\n\s*PCS", r"\1 PCS", table_text, flags=re.IGNORECASE
    )

    raw_lines = [
        line.strip() for line in table_text.splitlines() if line.strip()
    ]

    items = []
    current_item = None

    # 逐行狀態機掃描
    for line in raw_lines:
        # 跳過表格欄位名稱表頭
        if any(
            h in line.upper()
            for h in [
                "HIPPING MARK",
                "QUANTITY",
                "PART NUMBER",
                "PO NUMBER",
                "UNIT PRICE",
                "AMOUNT",
            ]
        ):
            continue

        # 遇到新的數量行 (例如 30 PCS, 120 PCS) -> 代表新商品開始
        qty_match = re.search(r"\b(\d+\s*PCS)\b", line, re.IGNORECASE)
        if qty_match and not any(k in line.upper() for k in ["USD", "CASE"]):
            # 結算並存入上一筆商品
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

        # 抓取品名規格 (包含直槓 | 且有英文，但不是 USD 金額行)
        if "|" in line and "USD" not in line.upper():
            desc_clean = line.replace("|", "").strip()
            # 排除純數字行的誤判
            if (
                desc_clean
                and not desc_clean.startswith("45")
                and not desc_clean.startswith("33")
            ):
                if not current_item["DESCRIPTION"]:
                    current_item["DESCRIPTION"] = desc_clean
                else:
                    current_item["DESCRIPTION"] += " " + desc_clean
            continue

        # 抓取 10 碼數字 (45 開頭分配給 PO NUMBER，33 或其他開頭分配給 PART NUMBER)
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

        # 抓取 USD 金額 (第 1 個通常為單價，第 2 個為總額)
        if "USD" in line.upper():
            prices = re.findall(r"USD\s*([0-9,\.]+)", line, re.IGNORECASE)
            for p in prices:
                clean_p = f"${clean_price(p)}"
                if not current_item["UNIT PRICE"]:
                    current_item["UNIT PRICE"] = clean_p
                elif not current_item["AMOUNT"]:
                    current_item["AMOUNT"] = clean_p
            continue

    # 結算最後一筆商品
    if current_item and current_item["QUANTITY"]:
        items.append(current_item)

    # 後處理：防呆檢核單價與總額是否顛倒
    for itm in items:
        try:
            up_num = float(itm["UNIT PRICE"].replace("$", ""))
            am_num = float(itm["AMOUNT"].replace("$", ""))
            if up_num > am_num and itm["QUANTITY"]:
                qty_digits = re.search(r"\d+", itm["QUANTITY"])
                if qty_digits and float(qty_digits.group(0)) > 1:
                    itm["UNIT PRICE"], itm["AMOUNT"] = (
                        itm["AMOUNT"],
                        itm["UNIT PRICE"],
                    )
        except Exception:
            pass

    return items


# ==================== 4. 網頁操作介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有 Excel 檔案")
    existing_file = st.file_uploader(
        "若要將新發票明細累加至現有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳上海騰飛 PDF 發票檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票（支援多個檔案同時上傳）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    all_records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        with st.spinner(f"正在執行第 {i+1} 份 PDF 的 OCR 掃描辨識與狀態機對齊..."):
            records = parse_tengfei_invoice(file.read())
            all_records.extend(records)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(all_records)

    # 累加合併 Excel
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

    # 結果預覽
    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    # 產出下載
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
