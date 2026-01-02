from pathlib import Path
import re

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data_pdf"

def roc_to_ad_ym(roc_year: int, month: int) -> str:
    ad_year = roc_year + 1911
    return f"{ad_year:04d}-{month:02d}"

def parse_ym_from_filename(name: str) -> str:
    """
    114年11月xxx.pdf -> 2025-11
    找不到就回傳空字串
    """
    m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", name)
    if not m:
        return ""
    return roc_to_ad_ym(int(m.group(1)), int(m.group(2)))

def normalize_filename(s: str) -> str:
    # 去掉半形/全形空白
    return s.replace(" ", "").replace("\u3000", "")
