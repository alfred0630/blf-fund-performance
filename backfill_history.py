import re
import time
import hashlib
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin

import requests
import pandas as pd
import pdfplumber
from bs4 import BeautifulSoup

# =========================
# 基本設定
# =========================
BASE_DIR = Path(__file__).resolve().parent
OUT_EXCEL = BASE_DIR / "blf_returns_backfill.xlsx"
DOWNLOAD_DIR = BASE_DIR / "data_pdf" / "_backfill"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
})

# ✅ 欄位順序與名稱：完全依照你要求的最新版本
COLS = [
    "年月",
    "勞保基金",
    "國民年金保險",
    "勞退新制",
    "勞退舊制",
    "退撫基金",                # 指 已實現收益率
    "退撫基金_整體收益率(%)",  # 指 整體收益率
    "TW10Y_YTD(%)",
    "TW_TAIEX_YTD(%)",
    "MSCI_World_YTD(%)",
    "BBG_GlobalBond_YTD(%)",
]

# =========================
# 共用小工具
# =========================
def roc_to_ad_ym(roc_year: int, month: int) -> str:
    return f"{roc_year + 1911:04d}-{month:02d}"

def clean_num(text):
    """清理數字字串，移除百分比、逗號、備註文字"""
    if not text: return 0.0
    cleaned = re.sub(r"[^\d\.-]", "", str(text))
    try:
        return float(cleaned) if cleaned else 0.0
    except:
        return 0.0

def parse_ym_from_filename(name: str) -> str:
    m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", name)
    if not m: return ""
    return roc_to_ad_ym(int(m.group(1)), int(m.group(2)))

def download_pdf(url: str, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    # 取檔名
    base = url.split("?")[0].split("/")[-1]
    if not base.lower().endswith(".pdf"):
        h = hashlib.md5(url.encode("utf-8")).hexdigest()[:8]
        filename = f"file_{h}.pdf"
    else:
        filename = base
    
    path = out_dir / filename
    if path.exists() and path.stat().st_size > 1000:
        return path

    r = SESSION.get(url, timeout=60)
    r.raise_for_status()
    path.write_bytes(r.content)
    return path

def upsert_row(df: pd.DataFrame, ym: str) -> int:
    if (df["年月"] == ym).any():
        return df.index[df["年月"] == ym][0]
    idx = len(df)
    df.loc[idx, "年月"] = ym
    return idx

def sort_by_ym(df: pd.DataFrame) -> pd.DataFrame:
    df["年月"] = pd.to_datetime(df["年月"], errors='coerce')
    df = df.dropna(subset=["年月"])
    df = df.sort_values("年月").reset_index(drop=True)
    df["年月"] = df["年月"].dt.strftime("%Y-%m")
    return df

# =========================
# 解析器 (全面更新)
# =========================

def parse_blf_retirement(pdf_path: Path) -> dict:
    ym = parse_ym_from_filename(pdf_path.name)
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:3]:
            for table in page.extract_tables() or []:
                for row in table or []:
                    if not row or len(row) < 3: continue
                    first = str(row[0])
                    if "收益率" in first or "報酬率" in first:
                        return {
                            "ym": ym,
                            "勞退新制": clean_num(row[1]),
                            "勞退舊制": clean_num(row[2])
                        }
    return {}

def parse_blf_labor_ins(pdf_path: Path) -> dict:
    ym = parse_ym_from_filename(pdf_path.name)
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:2]:
            for table in page.extract_tables() or []:
                header_idx = None
                for i, row in enumerate(table):
                    if "勞保基金" in "".join([str(c) for c in row]):
                        header_idx = i; break
                if header_idx is None: continue
                
                header = [str(c) for c in table[header_idx]]
                li_col = next((j for j, c in enumerate(header) if "勞保基金" in c), None)
                if li_col is None: continue
                
                for row in table[header_idx + 1:]:
                    if "收益率" in str(row[0]) or "報酬率" in str(row[0]):
                        return {"ym": ym, "勞保基金": clean_num(row[li_col])}
    return {}

def parse_blf_national_pension(pdf_path: Path) -> dict:
    ym = parse_ym_from_filename(pdf_path.name)
    if not ym: return {}
    labels = ["國民年金保險", "TW10Y_YTD(%)", "TW_TAIEX_YTD(%)", "MSCI_World_YTD(%)", "BBG_GlobalBond_YTD(%)"]
    
    with pdfplumber.open(str(pdf_path)) as pdf:
        for table in pdf.pages[0].extract_tables() or []:
            for row in table or []:
                if "收益率" in str(row[0]) or "報酬率" in str(row[0]):
                    nums = [clean_num(c) for c in row[1:6]]
                    if len(nums) == 5:
                        res = {"ym": ym}
                        res.update(dict(zip(labels, nums)))
                        return res
    return {}

def parse_fundgov_csf(pdf_path: Path) -> dict:
    """✅ 核心改動：爬退撫第二頁，抓已實現與整體"""
    with pdfplumber.open(str(pdf_path)) as pdf:
        if len(pdf.pages) < 2: return {}
        page = pdf.pages[1]
        table = page.extract_table()
        if not table: return {}

        target_row = None
        for row in reversed(table):
            # 正則確保抓到「115(1月)」這類格式，排除純註解
            if row and row[0] and re.search(r"\d+.*月", str(row[0])):
                target_row = row
                break
        
        if not target_row: return {}

        # 從行首文字解析年月
        m = re.search(r"(\d+)\s*\(\s*(\d+)\s*月\s*\)", str(target_row[0]))
        if m:
            ym = roc_to_ad_ym(int(m.group(1)), int(m.group(2)))
        else:
            return {}

        return {
            "ym": ym,
            "退撫基金": clean_num(target_row[2]),            # Index 2: 已實現
            "退撫基金_整體收益率(%)": clean_num(target_row[4])  # Index 4: 整體
        }

# =========================
# 抓取連結邏輯 (維持不變)
# =========================
def collect_blf_links(url):
    links = []
    for p in range(1, 15): # 抓前15頁通常就夠回補幾年了
        r = SESSION.get(f"{url}?Page={p}&PageSize=10")
        soup = BeautifulSoup(r.text, "html.parser")
        found = [urljoin(url, a["href"]) for a in soup.find_all("a", href=True) if ".pdf" in a["href"].lower() or "media/" in a["href"]]
        if not found: break
        links.extend(found)
    return list(set(links))

def collect_fundgov_links(url):
    links = []
    for p in range(1, 10):
        r = SESSION.get(f"{url}&page={p}")
        soup = BeautifulSoup(r.text, "html.parser")
        found = [urljoin(url, a["href"]) for a in soup.find_all("a", href=True) if "Download.ashx" in a["href"]]
        if not found: break
        links.extend(found)
    return list(set(links))

# =========================
# 主流程
# =========================
@dataclass
class FundJob:
    name: str; list_url: str; collect_fn: callable; parse_fn: callable; subdir: str

def main():
    jobs = [
        FundJob("勞退", "https://www.blf.gov.tw/49200/49255/49261/49269/49273/lpsimplelist", collect_blf_links, parse_blf_retirement, "retirement"),
        FundJob("勞保", "https://www.blf.gov.tw/49200/49255/49281/49285/49289/lpsimplelist", collect_blf_links, parse_blf_labor_ins, "labor_ins"),
        FundJob("國保", "https://www.blf.gov.tw/49200/49255/49323/49327/49331/lpsimplelist", collect_blf_links, parse_blf_national_pension, "national_pension"),
        FundJob("退撫", "https://www.fund.gov.tw/News.aspx?n=658&sms=11737", collect_fundgov_links, parse_fundgov_csf, "csf"),
    ]

    df = pd.DataFrame(columns=COLS)

    for job in jobs:
        print(f"\n🚀 開始回補：{job.name}")
        links = job.collect_fn(job.list_url)
        print(f"🔎 找到 {len(links)} 個連結")
        
        for i, url in enumerate(links, 1):
            try:
                path = download_pdf(url, DOWNLOAD_DIR / job.subdir)
                data = job.parse_fn(path)
                if data and "ym" in data:
                    idx = upsert_row(df, data["ym"])
                    for k, v in data.items():
                        if k in COLS: df.loc[idx, k] = v
                if i % 10 == 0: print(f"  進度: {i}/{len(links)}")
            except:
                continue
            time.sleep(0.1)

    # 排序與輸出
    df = sort_by_ym(df)
    # 確保欄位順序完全正確
    df = df[COLS]
    df.to_excel(OUT_EXCEL, index=False)
    print(f"\n🎉 回補完成！檔案存於: {OUT_EXCEL}")

if __name__ == "__main__":
    main()