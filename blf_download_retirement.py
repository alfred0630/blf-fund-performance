from urllib.parse import urljoin
from pathlib import Path
import requests
from bs4 import BeautifulSoup
from common import DATA_DIR

LIST_URL = "https://www.blf.gov.tw/49200/49255/49261/49269/49273/lpsimplelist"
OUT_DIR = DATA_DIR  # 勞退放在 data_pdf/ 根目錄也 OK

def pick_latest_pdf_link(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = (a.get_text() or "").strip()
        if ".pdf" in href.lower() or text.lower().endswith(".pdf"):
            return href
    raise ValueError("勞退：找不到 PDF 連結（網站結構可能改了）。")

def download_latest() -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    resp = requests.get(LIST_URL, timeout=30)
    resp.raise_for_status()

    href = pick_latest_pdf_link(resp.text)
    pdf_url = urljoin(LIST_URL, href)

    r = requests.get(pdf_url, timeout=30)
    r.raise_for_status()

    filename = pdf_url.split("?")[0].split("/")[-1]
    if not filename.lower().endswith(".pdf"):
        filename = "retirement_latest.pdf"

    out_path = OUT_DIR / filename
    out_path.write_bytes(r.content)

    print("✅ 勞退最新 PDF：", pdf_url)
    print("✅ 勞退已下載：", str(out_path))
    return out_path
