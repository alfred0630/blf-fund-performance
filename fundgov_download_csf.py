from urllib.parse import urljoin
from pathlib import Path
import requests
from bs4 import BeautifulSoup
from common import DATA_DIR

LIST_URL = "https://www.fund.gov.tw/News.aspx?n=658&sms=11737"
OUT_DIR = DATA_DIR / "civil_service_fund"

def pick_latest_pdf_link(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    pdfs = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if ".pdf" in href.lower() or "Download.ashx" in href:
            pdfs.append(href)
    if not pdfs:
        raise ValueError("退撫：找不到 PDF 連結。")
    return pdfs[0]

def download_latest() -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    resp = requests.get(LIST_URL, timeout=30)
    resp.raise_for_status()

    href = pick_latest_pdf_link(resp.text)
    pdf_url = urljoin(LIST_URL, href)

    r = requests.get(pdf_url, timeout=30)
    r.raise_for_status()

    # 這站常常拿不到乾淨檔名，所以就固定 latest.pdf
    out_path = OUT_DIR / "latest.pdf"
    out_path.write_bytes(r.content)

    print("✅ 退撫最新 PDF：", pdf_url)
    print("✅ 退撫已下載：", str(out_path))
    return out_path
