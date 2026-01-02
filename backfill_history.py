import re
import time
import hashlib
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

import requests
import pandas as pd
import pdfplumber
from bs4 import BeautifulSoup

# =========================
# 基本設定
# =========================
BASE_DIR = Path(__file__).resolve().parent
OUT_EXCEL = BASE_DIR / "blf_returns_backfill.xlsx"  # ✅ 輸出不同 Excel
DOWNLOAD_DIR = BASE_DIR / "data_pdf" / "_backfill"  # ✅ 歷史 PDF 都放這裡
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
})

# 欄位（跟你專業版一致）
COLS = [
    "年月",
    "新制基金_YTD(%)",
    "舊制基金_YTD(%)",
    "勞保基金_YTD(%)",
    "國保基金_YTD(%)",
    "TW10Y_YTD(%)",
    "TW_TAIEX_YTD(%)",
    "MSCI_World_YTD(%)",
    "BBG_GlobalBond_YTD(%)",
    "退撫基金_整體實際收益率(%)",
]

# =========================
# 共用小工具
# =========================
def roc_to_ad_ym(roc_year: int, month: int) -> str:
    return f"{roc_year + 1911:04d}-{month:02d}"

def parse_ym_from_filename(name: str) -> str:
    m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", name)
    if not m:
        return ""
    return roc_to_ad_ym(int(m.group(1)), int(m.group(2)))

def safe_filename_from_url(url: str) -> str:
    """
    從 URL 取出檔名；若取不到，就用 hash 當檔名避免衝突
    """
    base = url.split("?")[0].split("/")[-1]
    if base.lower().endswith(".pdf") and len(base) > 4:
        return base.replace(" ", "").replace("\u3000", "")
    h = hashlib.md5(url.encode("utf-8")).hexdigest()[:12]
    return f"{h}.pdf"

def download_pdf(url: str, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    filename = safe_filename_from_url(url)
    path = out_dir / filename
    if path.exists() and path.stat().st_size > 10_000:
        return path

    r = SESSION.get(url, timeout=60)
    r.raise_for_status()
    path.write_bytes(r.content)
    return path

def ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    for c in COLS:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def upsert_row(df: pd.DataFrame, ym: str) -> int:
    if (df["年月"] == ym).any():
        return df.index[df["年月"] == ym][0]
    idx = len(df)
    df.loc[idx, "年月"] = ym
    return idx

def sort_by_ym(df: pd.DataFrame) -> pd.DataFrame:
    df["年月"] = pd.to_datetime(df["年月"])
    df = df.sort_values("年月").reset_index(drop=True)
    df["年月"] = df["年月"].dt.strftime("%Y-%m")
    return df

# =========================
# 解析器（直接內建，不依賴你其他模組）
# =========================
def parse_blf_retirement(pdf_path: Path) -> dict:
    # 新制/舊制：抓表格「收益率/報酬率」那列的第 2/3 欄
    ym = parse_ym_from_filename(pdf_path.name)
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
                        if ym and m1 and m2:
                            return {
                                "ym": ym,
                                "新制基金_YTD(%)": float(m1.group(0)),
                                "舊制基金_YTD(%)": float(m2.group(0)),
                            }
    return {}

def parse_blf_labor_ins(pdf_path: Path) -> dict:
    # 勞保基金：表格中「勞保基金」欄，找「收益率/報酬率」列
    ym = parse_ym_from_filename(pdf_path.name)
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:2]:
            for table in page.extract_tables() or []:
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
                for row in table[header_idx + 1:]:
                    if not row or len(row) <= li_col:
                        continue
                    first = (row[0] or "").strip()
                    if "收益率" in first or "報酬率" in first:
                        val = (row[li_col] or "").replace("％", "%").strip()
                        m = re.search(r"-?\d+(?:\.\d+)?", val)
                        if ym and m:
                            return {"ym": ym, "勞保基金_YTD(%)": float(m.group(0))}
    return {}

def parse_blf_national_pension(pdf_path: Path) -> dict:
    # 國保：同一列「收益率/報酬率」後面 5 個百分比
    ym = parse_ym_from_filename(pdf_path.name)
    if not ym:
        return {}

    labels = [
        "國保基金_YTD(%)",
        "TW10Y_YTD(%)",
        "TW_TAIEX_YTD(%)",
        "MSCI_World_YTD(%)",
        "BBG_GlobalBond_YTD(%)",
    ]

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
    return {}

def parse_fundgov_csf(pdf_path: Path) -> dict:
    # 退撫：從 PDF 內文抓「xxx年截至xx月」當年月，並抓「整體」的「實際收益率(%)」
    with pdfplumber.open(str(pdf_path)) as pdf:
        page = pdf.pages[0]
        text = (page.extract_text() or "")
        compact = text.replace(" ", "").replace("\n", "")

        m = re.search(r"(\d{2,3})年截至(\d{1,2})月", compact)
        if not m:
            m = re.search(r"(\d{2,3})年.*?截至.*?(\d{1,2})月", compact)
        if not m:
            return {}
        ym = roc_to_ad_ym(int(m.group(1)), int(m.group(2)))

        # 抓「整體」那列的最後數字（實際收益率）
        mm = re.search(r"整\s*體\s*[-\d,\.]+\s+([-\d\.]+)\s*$", text, flags=re.MULTILINE)
        if mm:
            return {"ym": ym, "退撫基金_整體實際收益率(%)": float(mm.group(1))}

        # fallback：掃表格
        for table in (page.extract_tables() or []):
            for row in table or []:
                if not row:
                    continue
                first = (row[0] or "").strip()
                if "整" in first and "體" in first:
                    last = (row[-1] or "").replace("％", "%").strip()
                    mmm = re.search(r"-?\d+(?:\.\d+)?", last)
                    if mmm:
                        return {"ym": ym, "退撫基金_整體實際收益率(%)": float(mmm.group(0))}
    return {}

# =========================
# 抓取「歷史 PDF 連結」：BLF lpsimplelist（有 Page / PageSize）
# =========================
def collect_blf_pdf_links(list_url: str, max_pages: int = 200) -> list[str]:
    """
    BLF lpsimplelist：用 ?Page=1&PageSize=10 分頁一路抓到沒新連結為止
    """
    links = []
    seen = set()
    for page in range(1, max_pages + 1):
        url = f"{list_url}?Page={page}&PageSize=10"
        r = SESSION.get(url, timeout=30)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        page_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            # 站內常見 pdf 連結會含 media/ 或 .pdf
            if ".pdf" in href.lower() or "media/" in href.lower():
                full = urljoin(list_url, href)
                if full not in seen:
                    seen.add(full)
                    page_links.append(full)

        if not page_links:
            # 這頁沒抓到任何 PDF，通常表示已經到底
            break

        links.extend(page_links)

        # 友善一點，避免過快
        time.sleep(0.2)

    return links

# =========================
# 抓取「歷史 PDF 連結」：fund.gov.tw（頁面有很多頁）
# =========================
def collect_fundgov_pdf_links(list_url: str, max_pages: int = 200) -> list[str]:
    """
    fund.gov.tw News.aspx：嘗試用 &page=2 / &Page=2 取分頁
    會先抓第 1 頁的 page 連結推測總頁數，抓不到就跑到 max_pages 或無新連結為止
    """
    def fetch(url: str) -> str:
        r = SESSION.get(url, timeout=30)
        r.raise_for_status()
        return r.text

    html1 = fetch(list_url)
    soup1 = BeautifulSoup(html1, "html.parser")

    # 從頁碼連結抓最大 page（若有）
    max_found = 1
    for a in soup1.find_all("a", href=True):
        href = a["href"]
        m = re.search(r"[?&](?:page|Page)=(\d+)", href)
        if m:
            max_found = max(max_found, int(m.group(1)))
    total_pages = max_found if max_found > 1 else max_pages

    links = []
    seen = set()

    for p in range(1, total_pages + 1):
        if p == 1:
            html = html1
        else:
            # 先試 &page=
            url = list_url + ("" if "?" in list_url else "?")
            url = url + ("" if url.endswith("?") else "&")
            url_page = f"{url}page={p}"
            html = fetch(url_page)

            # 若這頁看起來不像換頁，就再試 &Page=
            if "Download.ashx" not in html and ".pdf" not in html.lower():
                url_page2 = f"{url}Page={p}"
                html = fetch(url_page2)

        soup = BeautifulSoup(html, "html.parser")
        page_links = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "Download.ashx" in href or ".pdf" in href.lower():
                full = urljoin(list_url, href)
                if full not in seen:
                    seen.add(full)
                    page_links.append(full)

        if not page_links and p > 3:
            # 後面連續沒東西就停（避免一直空轉）
            break

        links.extend(page_links)
        time.sleep(0.2)

    return links

# =========================
# 主流程：回補
# =========================
@dataclass
class FundJob:
    name: str
    list_url: str
    collect_links: callable
    parse_pdf: callable
    out_subdir: str

def main():
    jobs = [
        FundJob(
            name="勞退(新/舊制)",
            list_url="https://www.blf.gov.tw/49200/49255/49261/49269/49273/lpsimplelist",
            collect_links=collect_blf_pdf_links,
            parse_pdf=parse_blf_retirement,
            out_subdir="retirement"
        ),
        FundJob(
            name="勞保基金",
            list_url="https://www.blf.gov.tw/49200/49255/49281/49285/49289/lpsimplelist",
            collect_links=collect_blf_pdf_links,
            parse_pdf=parse_blf_labor_ins,
            out_subdir="labor_insurance"
        ),
        FundJob(
            name="國保基金(含4指標)",
            list_url="https://www.blf.gov.tw/49200/49255/49323/49327/49331/lpsimplelist",
            collect_links=collect_blf_pdf_links,
            parse_pdf=parse_blf_national_pension,
            out_subdir="national_pension"
        ),
        FundJob(
            name="退撫基金(整體實際收益率)",
            list_url="https://www.fund.gov.tw/News.aspx?n=658&sms=11737",
            collect_links=collect_fundgov_pdf_links,
            parse_pdf=parse_fundgov_csf,
            out_subdir="civil_service_fund"
        ),
    ]

    df = pd.DataFrame(columns=["年月"])
    df = ensure_cols(df)

    for job in jobs:
        print(f"\n=== 開始回補：{job.name} ===")
        links = job.collect_links(job.list_url)
        print(f"🔎 找到 PDF 連結數：{len(links)}")

        ok = 0
        fail = 0

        for i, url in enumerate(links, 1):
            try:
                pdf_path = download_pdf(url, DOWNLOAD_DIR / job.out_subdir)
                payload = job.parse_pdf(pdf_path)

                if not payload:
                    fail += 1
                    continue

                ym = payload.get("ym")
                if not ym:
                    fail += 1
                    continue

                idx = upsert_row(df, ym)
                for k, v in payload.items():
                    if k == "ym":
                        continue
                    if k not in df.columns:
                        df[k] = pd.NA
                    df.loc[idx, k] = v

                ok += 1

            except Exception as e:
                fail += 1

            # 每 20 個印一次進度
            if i % 20 == 0:
                print(f"  進度 {i}/{len(links)} | 成功 {ok} | 失敗 {fail}")

            time.sleep(0.15)

        print(f"✅ {job.name} 完成：成功 {ok}，失敗 {fail}")

    df = ensure_cols(df)
    df = sort_by_ym(df)

    # 寫出新的 Excel（欄位名稱一致，但檔名不同）
    with pd.ExcelWriter(OUT_EXCEL, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="data", index=False)

    print(f"\n🎉 回補完成！輸出：{OUT_EXCEL}")

if __name__ == "__main__":
    main()
