import re
from pathlib import Path
import pdfplumber
from common import roc_to_ad_ym


def parse_ym_from_row(raw_text: str) -> str | None:
    """
    從第一欄解析年月，支援：
    - 114(11 月) 註2 -> 2025-11
    - 114(11月)     -> 2025-11
    - 114           -> 2025-12   # 年底資料視為12月
    - 113(註1)      -> 2024-12
    """
    if not raw_text:
        return None

    text = str(raw_text).strip()

    # 先抓「年(月)」格式
    m = re.search(r"(\d+)\s*\(\s*(\d+)\s*月\s*\)", text)
    if m:
        roc_year = int(m.group(1))
        month = int(m.group(2))
        return roc_to_ad_ym(roc_year, month)

    # 再抓只有年份的格式，視為該年12月
    # 例如：114、113(註1)
    m = re.match(r"^\s*(\d+)", text)
    if m:
        roc_year = int(m.group(1))
        return roc_to_ad_ym(roc_year, 12)

    return None


def is_valid_perf_row(row) -> bool:
    """
    判斷是否為退撫基金歷年績效表的有效資料列。
    只要第一欄能成功解析年月，就視為有效資料列。
    """
    if not row or len(row) < 5:
        return False

    first_col = row[0]
    if not first_col:
        return False

    return parse_ym_from_row(first_col) is not None


def clean_num(text) -> float:
    """
    清理數字字串，只保留數字、小數點、負號。
    範例：
    - '13.42%'   -> 13.42
    - '-2.31 %'  -> -2.31
    - '1,234.56' -> 1234.56
    - '' / None  -> 0.0
    """
    if text is None:
        return 0.0

    cleaned = re.sub(r"[^\d.\-]", "", str(text))
    return float(cleaned) if cleaned else 0.0


def extract(pdf_path: Path) -> dict:
    with pdfplumber.open(str(pdf_path)) as pdf:
        if len(pdf.pages) < 2:
            raise ValueError("退撫：PDF 頁數不足，找不到第二頁績效表")

        # 第二頁（index = 1）
        page = pdf.pages[1]
        table = page.extract_table()

        if not table:
            raise ValueError("退撫：在第二頁找不到表格內容")

        # 從底部往上找最新的一筆有效資料列
        target_row = None
        for row in reversed(table):
            if is_valid_perf_row(row):
                target_row = row
                break

        if not target_row:
            raise ValueError("退撫：找不到有效的歷年績效資料列")

        ym = parse_ym_from_row(target_row[0])
        if not ym:
            raise ValueError(f"退撫：日期格式解析失敗，原始內容：{target_row[0]}")

        try:
            # 表格欄位通常為：
            # [0] 年度
            # [1] 已實現收益數(億元)
            # [2] 已實現收益率(%)
            # [3] 整體收益數(億元)
            # [4] 整體收益率(%)
            realized_rate = clean_num(target_row[2])
            overall_rate = clean_num(target_row[4])
        except Exception as e:
            raise ValueError(f"退撫：數值解析失敗。內容：{target_row}，錯誤：{e}")

    return {
        "ym": ym,
        "csf_realized_return": realized_rate,
        "csf_overall_return": overall_rate,
    }