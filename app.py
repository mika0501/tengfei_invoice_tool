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
