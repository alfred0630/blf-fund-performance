import re
from pathlib import Path
import pdfplumber
from common import roc_to_ad_ym

def parse_ym_from_row(raw_text: str) -> str:
    """
    從行首文字如 '115(1 月) 註2' 中精確過濾出 YYYY-MM
    """
    if not raw_text:
        return None
    # 使用正則表達式只抓取數字部分：(\d+)年 (\d+)月
    # 這樣即便有 "註2" 或換行符號也會被無視
    m = re.search(r"(\d+)\s*\(\s*(\d+)\s*月\s*\)", str(raw_text))
    if m:
        roc_year = int(m.group(1))
        month = int(m.group(2))
        return roc_to_ad_ym(roc_year, month)
    return None

def extract(pdf_path: Path) -> dict:
    with pdfplumber.open(str(pdf_path)) as pdf:
        if len(pdf.pages) < 2:
            raise ValueError("退撫：PDF 頁數不足，找不到第二頁績效表")
        
        page = pdf.pages[1]
        table = page.extract_table()
        if not table:
            raise ValueError("退撫：在第二頁找不到表格內容")

        target_row = None
        for row in reversed(table):
            # 確保第一欄位包含數字且有「月」字眼，避免抓到純註解行
            if row and row[0] and re.search(r"\d+.*月", str(row[0])):
                target_row = row
                break

        if not target_row:
            raise ValueError("退撫：找不到有效的日期數據行")

        # 🟢 清洗日期：確保只回傳 "2026-01" 這種乾淨格式
        ym = parse_ym_from_row(target_row[0])
        if not ym:
            raise ValueError(f"退撫：日期格式解析失敗，原始內容：{target_row[0]}")

        # 抓取數值
        try:
            # 去除所有非數字與小數點的字元（如 ％, 註, 逗號）
            def clean_num(text):
                if not text: return 0.0
                cleaned = re.sub(r"[^\d\.-]", "", str(text))
                return float(cleaned) if cleaned else 0.0

            realized = clean_num(target_row[2])
            overall = clean_num(target_row[4])
        except Exception as e:
            raise ValueError(f"退撫：數值解析失敗。內容：{target_row}，錯誤：{e}")

    return {
        "ym": ym,
        "csf_realized_return": realized,
        "csf_overall_return": overall
    }