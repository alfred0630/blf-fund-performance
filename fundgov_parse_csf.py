import re
from pathlib import Path
import pdfplumber
from common import roc_to_ad_ym

def parse_ym_from_pdf(pdf_path: Path) -> str:
    with pdfplumber.open(str(pdf_path)) as pdf:
        text = (pdf.pages[0].extract_text() or "")
    compact = text.replace(" ", "").replace("\n", "")
    m = re.search(r"(\d{2,3})年截至(\d{1,2})月", compact)
    if not m:
        m = re.search(r"(\d{2,3})年.*?截至.*?(\d{1,2})月", compact)
    if not m:
        raise ValueError("退撫：PDF 內文找不到『xxx年截至xx月』")
    return roc_to_ad_ym(int(m.group(1)), int(m.group(2)))

def extract_overall_actual_return(pdf_path: Path) -> float:
    """
    抓『整體』的『實際收益率(%)』
    """
    with pdfplumber.open(str(pdf_path)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        m = re.search(r"整\s*體\s*[-\d,\.]+\s+([-\d\.]+)\s*$", text, flags=re.MULTILINE)
        if m:
            return float(m.group(1))

        for table in (page.extract_tables() or []):
            for row in table or []:
                if not row:
                    continue
                first = (row[0] or "").strip()
                if "整" in first and "體" in first:
                    last = (row[-1] or "").replace("％", "%").strip()
                    mm = re.search(r"-?\d+(?:\.\d+)?", last)
                    if mm:
                        return float(mm.group(0))

    raise ValueError("退撫：找不到『整體』的『實際收益率(%)』")

def extract(pdf_path: Path) -> dict:
    ym = parse_ym_from_pdf(pdf_path)
    csf = extract_overall_actual_return(pdf_path)
    return {"ym": ym, "csf_overall_actual_return": csf}
