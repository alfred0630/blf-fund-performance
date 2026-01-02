import re
from pathlib import Path
import pdfplumber
from common import parse_ym_from_filename

def extract(pdf_path: Path) -> dict:
    """
    回傳：
    {
      "ym": "2025-11",
      "npi_ytd": 13.34,
      "tw10y_ytd": 1.48,
      "taiex_ytd": 19.93,
      "msci_world_ytd": 21.07,
      "bbg_globalbond_ytd": 7.89
    }
    """
    ym = parse_ym_from_filename(pdf_path.name)
    labels = ["npi_ytd","tw10y_ytd","taiex_ytd","msci_world_ytd","bbg_globalbond_ytd"]

    with pdfplumber.open(str(pdf_path)) as pdf:
        text = pdf.pages[0].extract_text() or ""
        m = re.search(
            r"收益率/報酬率\s*([-\d\.]+)\s*％\s*([-\d\.]+)\s*％\s*([-\d\.]+)\s*％\s*([-\d\.]+)\s*％\s*([-\d\.]+)\s*％",
            text
        )
        if m:
            nums = list(map(float, m.groups()))
            out = {"ym": ym}
            out.update(dict(zip(labels, nums)))
            return out

        # fallback：掃表格
        for table in pdf.pages[0].extract_tables() or []:
            for row in table or []:
                if not row:
                    continue
                first = (row[0] or "").strip()
                if "收益率" in first or "報酬率" in first:
                    nums = []
                    for cell in row[1:6]:
                        cell = (cell or "").replace("％", "%").strip()
                        mm = re.search(r"-?\d+(?:\.\d+)?", cell)
                        if mm:
                            nums.append(float(mm.group(0)))
                    if len(nums) == 5:
                        out = {"ym": ym}
                        out.update(dict(zip(labels, nums)))
                        return out

    raise ValueError("國保：找不到『收益率/報酬率』5 欄。")
