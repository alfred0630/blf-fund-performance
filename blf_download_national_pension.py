from urllib.parse import urljoin
from pathlib import Path
import requests
from bs4 import BeautifulSoup
from common import DATA_DIR

LIST_URL = "https://www.blf.gov.tw/49200/49255/49323/49327/49331/lpsimplelist"
OUT_DIR = DATA_DIR / "national_pension"

def pick_latest_pdf_link(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = (a.get_text() or "").strip()
        if ".pdf" in href.lower() or text.lower().endswith(".pdf"):
            return href
    raise ValueError("國保：找不到 PDF 連結。")

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
        filename = "national_pension_latest.pdf"

    out_path = OUT_DIR / filename
    out_path.write_bytes(r.content)

    print("✅ 國保最新 PDF：", pdf_url)
    print("✅ 國保已下載：", str(out_path))
    return out_path
