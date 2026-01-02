import re
from pathlib import Path
import pdfplumber
from common import parse_ym_from_filename

def extract(pdf_path: Path) -> dict:
    """
    回傳：
    {
      "ym": "2025-11",
      "retirement_new_ytd": 13.70,
      "retirement_old_ytd": 20.07
    }
    """
    ym = parse_ym_from_filename(pdf_path.name)
    if not ym:
        # 理論上勞退檔名有 114年xx月，沒有就先空
        ym = ""

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:3]:
            for table in page.extract_tables() or []:
                for row in table or []:
                    if not row or len(row) < 3:
                        continue
                    first = (row[0] or "").strip()
                    if "收益率" in first or "報酬率" in first:
                        new = (row[1] or "").replace("％", "%").strip()
                        old = (row[2] or "").replace("％", "%").strip()
                        m1 = re.search(r"-?\d+(?:\.\d+)?", new)
                        m2 = re.search(r"-?\d+(?:\.\d+)?", old)
                        if m1 and m2:
                            return {
                                "ym": ym,
                                "retirement_new_ytd": float(m1.group(0)),
                                "retirement_old_ytd": float(m2.group(0)),
                            }
    raise ValueError("勞退：找不到『收益率/報酬率』。")
