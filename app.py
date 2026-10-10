import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="上海騰飛發票表格轉換工具", page_icon="📋", layout="wide"
)
st.title("📋 上海騰飛 (Tengfei) 發票表格自動轉 Excel 工具")
st.caption("已升級彈性切分演算法：完整抓取所有明細品項並轉為標準 Excel！")


def parse_tengfei_invoice(file_bytes):
    """解析上海騰飛發票檔案"""
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

    # ================= 2. 表格明細區塊拆解 =================
    # 截取表格區域
    table_block_match = re.search(
        r"HIPPING\s*MARK[\s\S]*?(?=COUNTRY\s*of|SUBTOTAL|THIS\s*TOTAL)",
        full_text,
        re.IGNORECASE,
    )
    table_text = table_block_match.group(0) if table_block_match else full_text

    # 依照金額標記 "| USD" 或換行模式做彈性切割
    # 每個項目通常結尾於第二個 "| USD xxx" (AMOUNT)
    # 我們利用正則切分出每一個獨立明細段落
    pattern = re.compile(
        r"(?P<qty>\d+[\s\n]*PCS)[\s\S]*?\|\s*(?P<desc>[A-Z0-9\s\.\*'\-,]+?)[\s\n]+(?P<codes>(?:\d{10}[\s\n]*)+)[\s\S]*?\|\s*USD\s*(?P<price>[0-9,\.]+)[\s\S]*?\|\s*USD\s*(?P<amt>[0-9,\.]+)",
        re.IGNORECASE,
    )

    items = []
    for match in pattern.finditer(table_text):
        # 處理數量
        qty_raw = match.group("qty")
        qty = re.sub(r"\s+", " ", qty_raw).strip()

        # 處理描述
        desc = match.group("desc").strip().replace("\n", " ")

        # 處理 10 碼數字 (PO Number 通常是 45 開頭, Part Number 通常是 33 開頭)
        codes_text = match.group("codes")
        found_codes = re.findall(r"\b\d{10}\b", codes_text)

        po_number = ""
        part_number = ""
        for c in found_codes:
            if c.startswith("45"):
                po_number = c
            else:
                part_number = c

        # 若沒特別分開，依順序帶入
        if not po_number and len(found_codes) >= 2:
            po_number = found_codes[1]
            part_number = found_codes[0]
        elif not po_number and len(found_codes) == 1:
            if found_codes[0].startswith("45"):
                po_number = found_codes[0]
            else:
                part_number = found_codes[0]

        # 處理金額
        unit_price = f"${match.group('price').strip().replace(',', '.')}"
        amount = f"${match.group('amt').strip().replace(',', '.')}"

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

    # 備援演算法：如果正則未抓滿，啟用逐行狀態機備援
    if len(items) == 0:
        lines = [line.strip() for line in table_text.splitlines() if line.strip()]
        cur_qty = ""
        cur_desc = ""
        cur_codes = []
        cur_prices = []

        for line in lines:
            # 檢查數量
            qty_chk = re.match(r"^(\d+\s*PCS)$", line, re.IGNORECASE)
            if qty_chk:
                # 若已有上一筆，先結算
                if cur_qty and (cur_codes or cur_prices):
                    po = next((c for c in cur_codes if c.startswith("45")), cur_codes[1] if len(cur_codes) > 1 else "")
                    pn = next((c for c in cur_codes if not c.startswith("45")), cur_codes[0] if cur_codes else "")
                    up = f"${cur_prices[0]}" if len(cur_prices) > 0 else ""
                    am = f"${cur_prices[1]}" if len(cur_prices) > 1 else ""
                    items.append({
                        "INVOICE NUMBER": invoice_number,
                        "INVOICE DATE": invoice_date,
                        "QUANTITY": cur_qty,
                        "PART NUMBER": pn,
                        "PO NUMBER": po,
                        "DESCRIPTION": cur_desc,
                        "UNIT PRICE": up,
                        "AMOUNT": am,
                    })
                cur_qty = qty_chk.group(1)
                cur_desc = ""
                cur_codes = []
                cur_prices = []
                continue

            if "|" in line and "USD" not in line:
                cur_desc = line.replace("|", "").strip()
            
            # 抓取 10 碼數字
            codes = re.findall(r"\b\d{10}\b", line)
            if codes:
                cur_codes.extend(codes)

            # 抓取 USD 金額
            if "USD" in line:
                prices = re.findall(r"USD\s*([0-9,\.]+)", line)
                for p in prices:
                    cur_prices.append(p.replace(",", "."))

        # 結算最後一筆
        if cur_qty:
            po = next((c for c in cur_codes if c.startswith("45")), cur_codes[1] if len(cur_codes) > 1 else "")
            pn = next((c for c in cur_codes if not c.startswith("45")), cur_codes[0] if cur_codes else "")
            up = f"${cur_prices[0]}" if len(cur_prices) > 0 else ""
            am = f"${cur_prices[1]}" if len(cur_prices) > 1 else ""
            items.append({
                "INVOICE NUMBER": invoice_number,
                "INVOICE DATE": invoice_date,
                "QUANTITY": cur_qty,
                "PART NUMBER": pn,
                "PO NUMBER": po,
                "DESCRIPTION": cur_desc,
                "UNIT PRICE": up,
                "AMOUNT": am,
            })

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
