import re
from pathlib import Path
import pdfplumber
from common import parse_ym_from_filename

def extract(pdf_path: Path) -> dict:
    """
    回傳：
    {"ym": "2025-11", "labor_ins_ytd": 13.41}
    """
    ym = parse_ym_from_filename(pdf_path.name)

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:2]:
            tables = page.extract_tables() or []
            for table in tables:
                if not table:
                    continue

                # 找 header 含「勞保基金」
                header_idx = None
                for i, row in enumerate(table):
                    row_text = " ".join([(c or "").strip() for c in row])
                    if "勞保基金" in row_text:
                        header_idx = i
                        break
                if header_idx is None:
                    continue

                header = [(c or "").strip() for c in table[header_idx]]
                li_col = next((j for j, c in enumerate(header) if "勞保基金" in c), None)
                if li_col is None:
                    continue

                for row in table[header_idx + 1 :]:
                    if not row or len(row) <= li_col:
                        continue
                    first = (row[0] or "").strip()
                    if "收益率" in first or "報酬率" in first:
                        val = (row[li_col] or "").replace("％", "%").strip()
                        m = re.search(r"-?\d+(?:\.\d+)?", val)
                        if m:
                            return {"ym": ym, "labor_ins_ytd": float(m.group(0))}
    raise ValueError("勞保：找不到『收益率/報酬率』。")
