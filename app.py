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

st.title("📋 上海騰飛發票表格自動轉 Excel 工具（精確防呆版）")
st.caption(
    "結合 300 DPI 影像增強 OCR 掃描 ＋ 逐行狀態機，完整容錯折行與黏字！"
)

# ==================== 1. 檔案上傳區塊（置頂醒目顯示） ====================
st.markdown("---")
col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    st.markdown("### 📥 步驟 1：上傳 PDF 發票檔案（必填）")
    uploaded_files = st.file_uploader(
        "請點擊瀏覽或將 PDF 檔案直接拖曳至此區域（可選多個檔案）",
        type=["pdf"],
        accept_multiple_files=True,
        key="main_pdf_uploader",
    )

with col_upload2:
    st.markdown("### 📑 步驟 2：上傳現有 Excel 檔案（選填，累加用）")
    existing_file = st.file_uploader(
        "若要將本次辨識的明細接續累加至既有的 Excel 表格，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel_uploader",
    )
st.markdown("---")


# ==================== 2. 資料清洗與 OCR 核心 ====================
def clean_price(val_str):
    """清理金額字串中的逗號與雜訊"""
    if not val_str:
        return ""
    val = val_str.replace(" ", "").strip()
    if "," in val and "." in val:
        val = val.replace(",", "")
    else:
        val = val.replace(",", ".")
    return val


def extract_text_with_ocr(file_bytes):
    """提取文字：若數位層文字不足或格式破碎，強制走 300 DPI OCR 影像辨識"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            # 若文字少於 400 字元，表示缺少完整文字層，強制啟動高解析度 OCR
            if len(txt.strip()) < 400:
                try:
                    img = page.to_image(resolution=300).original
                    img = img.convert("L")
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(2.2)
                    ocr_res = pytesseract.image_to_string(img, lang="eng")
                    if ocr_res and len(ocr_res.strip()) > len(txt.strip()):
                        txt = ocr_res
                except Exception as e:
                    st.warning(f"OCR 影像掃描提示: {e}")
            full_text += txt + "\n"
    return full_text


def parse_tengfei_invoice(file_bytes):
    """解析發票內容並進行狀態機對齊"""
    full_text = extract_text_with_ocr(file_bytes)

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

    # 2. 定位商品表格核心區域
    table_block_match = re.search(
        r"HIPPING\s*MARK[\s\S]*?(?=COUNTRY\s*of|SUBTOTAL|THIS\s*TOTAL)",
        full_text,
        re.IGNORECASE,
    )
    table_text = table_block_match.group(0) if table_block_match else full_text

    # 修正折行或黏在一起的 PCS (例如 30\nPCS 轉 30 PCS，1120PCS 轉 1120 PCS)
    table_text = re.sub(
        r"(\d+)\s*\n\s*PCS", r"\1 PCS", table_text, flags=re.IGNORECASE
    )
    table_text = re.sub(
        r"(\d+)(PCS|PCE)", r"\1 \2", table_text, flags=re.IGNORECASE
    )

    raw_lines = [
        line.strip() for line in table_text.splitlines() if line.strip()
    ]

    items = []
    current_item = None

    for line in raw_lines:
        # 排除表頭行
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

        # 遇到數量行判定為新項目起點 (例如 30 PCS, 1120 PCS)
        qty_match = re.search(
            r"\b(\d+)\s*(PCS|PCE)\b", line, re.IGNORECASE
        )
        if qty_match and not any(k in line.upper() for k in ["USD", "CASE"]):
            if current_item and current_item["QUANTITY"]:
                items.append(current_item)
            current_item = {
                "INVOICE NUMBER": invoice_number,
                "INVOICE DATE": invoice_date,
                "QUANTITY": f"{qty_match.group(1)} {qty_match.group(2).upper()}",
                "PART NUMBER": "",
                "PO NUMBER": "",
                "DESCRIPTION": "",
                "UNIT PRICE": "",
                "AMOUNT": "",
            }
            continue

        if not current_item:
            continue

        # 品名規格：包含直槓 | 且不是金額行
        if "|" in line and "USD" not in line.upper():
            desc_clean = line.replace("|", "").strip()
            # 排除純代碼數字
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

        # 10 碼代碼自動分流 (45 開頭分配給 PO，33 或其他開頭分配給 Part No)
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

        # USD 金額
        if "USD" in line.upper():
            prices = re.findall(r"USD\s*([0-9,\.]+)", line, re.IGNORECASE)
            for p in prices:
                clean_p = f"${clean_price(p)}"
                if not current_item["UNIT PRICE"]:
                    current_item["UNIT PRICE"] = clean_p
                elif not current_item["AMOUNT"]:
                    current_item["AMOUNT"] = clean_p
            continue

    if current_item and current_item["QUANTITY"]:
        items.append(current_item)

    # 單價與金額反置校驗
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


# ==================== 3. 處理與輸出結果 ====================
if uploaded_files:
    all_records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        with st.spinner(f"正在對「{file.name}」進行高解析 OCR 辨識與狀態機對齊..."):
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
