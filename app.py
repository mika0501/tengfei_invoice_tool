import io
import os
import re
import pandas as pd
import pdfplumber
from PIL import Image, ImageEnhance
import pytesseract
import streamlit as st

# Windows 本機環境備援路徑
if os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )

st.set_page_config(
    page_title="上海騰飛發票 OCR 表格轉換工具",
    page_icon="📋",
    layout="wide",
)
st.title("📋 上海騰飛發票紙本掃描轉 Excel 工具（OCR 辨識版）")
st.caption(
    "支援純紙本掃描 PDF，先以 300 DPI 影像增強去除雜訊，再透過 OCR 逐筆解析商品明細！"
)


def clean_num_str(val_str):
    """清理金額字串中的噪點與逗號"""
    if not val_str:
        return ""
    val = val_str.replace(" ", "").strip()
    if "," in val and "." in val:
        val = val.replace(",", "")
    else:
        val = val.replace(",", ".")
    return val


def extract_text_via_ocr(file_bytes):
    """將 PDF 每一頁轉為圖片並執行 OCR 辨識"""
    extracted_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            # 先嘗試直接提取文字
            page_text = page.extract_text() or ""

            # 若直接提取文字很少（小於 50 字元，表示為掃描圖），執行高解析度 OCR
            if len(page_text.strip()) < 50:
                try:
                    img = page.to_image(resolution=300).original
                    # 影像前處理：轉灰階並提高對比度，去除傳真噪點
                    img = img.convert("L")
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(2.2)
                    ocr_res = pytesseract.image_to_string(img, lang="eng")
                    if ocr_res:
                        page_text = ocr_res
                except Exception as e:
                    st.warning(f"OCR 處理頁面時發生提示: {e}")

            extracted_text += page_text + "\n"

    return extracted_text


def parse_tengfei_ocr(full_text):
    """解析 OCR 辨識後的文字內容"""
    # 1. 表頭核心欄位
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

    # 2. 表格區塊處理
    # 統一換行的 PCS（如 30\nPCS 轉為 30 PCS）
    text_fixed = re.sub(
        r"(\d+)\s*\n\s*PCS", r"\1 PCS", full_text, flags=re.IGNORECASE
    )

    # 以「數字 PCS」作為商品項目的錨點分割
    split_indices = [
        m.start()
        for m in re.finditer(r"\b\d+\s*PCS\b", text_fixed, re.IGNORECASE)
    ]
    split_indices.append(len(text_fixed))

    items = []

    for i in range(len(split_indices) - 1):
        block = text_fixed[split_indices[i] : split_indices[i + 1]].strip()
        if not block:
            continue

        # 數量
        qty_match = re.search(r"\b(\d+\s*PCS)\b", block, re.IGNORECASE)
        qty = qty_match.group(1).strip() if qty_match else ""

        # 10 碼數字（45 開頭判定為 PO，其餘為 Part No）
        codes = re.findall(r"\b\d{10}\b", block)
        po_number = ""
        part_number = ""
        for c in codes:
            if c.startswith("45"):
                po_number = c
            elif c.startswith("33"):
                part_number = c

        if not po_number and len(codes) >= 2:
            po_number = codes[1]
        if not part_number and len(codes) >= 1:
            part_number = codes[0]

        # 品名與規格描述
        desc_match = re.search(r"\|\s*([A-Za-z0-9\s\.\*'\-,]+?)\n", block)
        if desc_match:
            desc = desc_match.group(1).strip()
        else:
            desc_line = re.search(r"(CASE\s*BRACKET[^\n\r\|]+)", block)
            desc = desc_line.group(1).strip() if desc_line else ""

        # 單價與金額
        prices = re.findall(r"USD\s*([0-9,\.]+)", block, re.IGNORECASE)
        unit_price = ""
        amount = ""
        if len(prices) >= 2:
            unit_price = f"${clean_num_str(prices[0])}"
            amount = f"${clean_num_str(prices[1])}"
        elif len(prices) == 1:
            amount = f"${clean_num_str(prices[0])}"

        if qty and (codes or prices or desc):
            items.append(
                {
                    "INVOICE NUMBER": invoice_number,
                    "INVOICE DATE": invoice_date,
                    "QUANTITY": qty,
                    "PART NUMBER": part_number,
                    "PO NUMBER": po_number,
                    "DESCRIPTION": desc,
                    "UNIT PRICE": unit_price,
                    "AMOUNT": amount,
                }
            )

    return items


# ==================== 網頁介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有 Excel 檔案")
    existing_file = st.file_uploader(
        "若要將明細接續累加至既有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳紙本掃描 PDF")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票（支援多個檔案）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    all_records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        with st.spinner(f"正在進行第 {i+1} 份 PDF 的 OCR 光學掃描辨識中..."):
            extracted_text = extract_text_via_ocr(file.read())
            records = parse_tengfei_ocr(extracted_text)
            all_records.extend(records)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(all_records)

    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            st.success(
                f"🎉 OCR 成功辨識 {len(new_df)} 筆明細，並已合併至原有的 Excel！目前累積共 {len(final_df)} 筆。"
            )
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 OCR 成功辨識 {len(new_df)} 筆商品明細資料！")

    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Invoice_Items")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載轉換後的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="Tengfei_OCR發票明細.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
