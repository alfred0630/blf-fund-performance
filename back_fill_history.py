import re
import sys
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
OUT_EXCEL = BASE_DIR / "blf_returns_backfill(temporary).xlsx"
DOWNLOAD_DIR = BASE_DIR / "data_pdf" / "_backfill"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
})

COLS = [
    "年月",
    "勞保基金",
    "國民年金保險",
    "勞退新制",
    "勞退舊制",
    "退撫基金",                # 已實現收益率
    "退撫基金_整體收益率(%)",  # 整體收益率
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
    if text is None:
        return 0.0
    cleaned = re.sub(r"[^\d\.-]", "", str(text))
    try:
        return float(cleaned) if cleaned else 0.0
    except Exception:
        return 0.0


def parse_ym_from_filename(name: str) -> str:
    m = re.search(r"(\d{2,3})\s*年\s*(\d{1,2})\s*月", name)
    if not m:
        return ""
    return roc_to_ad_ym(int(m.group(1)), int(m.group(2)))


def validate_ym_format(ym: str) -> str:
    if not re.fullmatch(r"\d{4}-\d{2}", ym):
        raise ValueError(f"start_ym 格式錯誤：{ym}，應為 YYYY-MM")
    year, month = ym.split("-")
    month = int(month)
    if not (1 <= month <= 12):
        raise ValueError(f"start_ym 月份錯誤：{ym}")
    return ym


def ym_to_int(ym: str) -> int:
    year, month = ym.split("-")
    return int(year) * 100 + int(month)


def should_keep_ym(ym: str, start_ym: str | None) -> bool:
    if not ym:
        return False
    if start_ym is None:
        return True
    return ym_to_int(ym) >= ym_to_int(start_ym)


def download_pdf(url: str, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)

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
    df["年月"] = pd.to_datetime(df["年月"], errors="coerce")
    df = df.dropna(subset=["年月"])
    df = df.sort_values("年月").reset_index(drop=True)
    df["年月"] = df["年月"].dt.strftime("%Y-%m")
    return df


# =========================
# 解析器
# =========================
def parse_blf_retirement(pdf_path: Path) -> dict:
    ym = parse_ym_from_filename(pdf_path.name)
    if not ym:
        return {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:3]:
            for table in page.extract_tables() or []:
                for row in table or []:
                    if not row or len(row) < 3:
                        continue
                    first = str(row[0])
                    if "收益率" in first or "報酬率" in first:
                        return {
                            "ym": ym,
                            "勞退新制": clean_num(row[1]),
                            "勞退舊制": clean_num(row[2]),
                        }
    return {}


def parse_blf_labor_ins(pdf_path: Path) -> dict:
    ym = parse_ym_from_filename(pdf_path.name)
    if not ym:
        return {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[:2]:
            for table in page.extract_tables() or []:
                header_idx = None
                for i, row in enumerate(table):
                    joined = "".join([str(c) for c in row if c is not None])
                    if "勞保基金" in joined:
                        header_idx = i
                        break

                if header_idx is None:
                    continue

                header = [str(c) for c in table[header_idx]]
                li_col = next((j for j, c in enumerate(header) if "勞保基金" in c), None)
                if li_col is None:
                    continue

                for row in table[header_idx + 1:]:
                    if not row:
                        continue
                    first = str(row[0])
                    if "收益率" in first or "報酬率" in first:
                        return {
                            "ym": ym,
                            "勞保基金": clean_num(row[li_col]),
                        }
    return {}


def parse_blf_national_pension(pdf_path: Path) -> dict:
    ym = parse_ym_from_filename(pdf_path.name)
    if not ym:
        return {}

    labels = [
        "國民年金保險",
        "TW10Y_YTD(%)",
        "TW_TAIEX_YTD(%)",
        "MSCI_World_YTD(%)",
        "BBG_GlobalBond_YTD(%)",
    ]

    with pdfplumber.open(str(pdf_path)) as pdf:
        for table in pdf.pages[0].extract_tables() or []:
            for row in table or []:
                if not row:
                    continue
                first = str(row[0])
                if "收益率" in first or "報酬率" in first:
                    nums = [clean_num(c) for c in row[1:6]]
                    if len(nums) == 5:
                        res = {"ym": ym}
                        res.update(dict(zip(labels, nums)))
                        return res
    return {}


def parse_csf_ym_from_row(raw_text: str) -> str | None:
    """
    支援：
    - 114(11 月) 註2 -> 2025-11
    - 114(11月)     -> 2025-11
    - 114           -> 2025-12
    - 113(註1)      -> 2024-12
    """
    if not raw_text:
        return None

    text = str(raw_text).strip()

    m = re.search(r"(\d+)\s*\(\s*(\d+)\s*月\s*\)", text)
    if m:
        roc_year = int(m.group(1))
        month = int(m.group(2))
        return roc_to_ad_ym(roc_year, month)

    m = re.match(r"^\s*(\d+)", text)
    if m:
        roc_year = int(m.group(1))
        return roc_to_ad_ym(roc_year, 12)

    return None


def is_valid_csf_perf_row(row) -> bool:
    if not row or len(row) < 5:
        return False

    first_col = row[0]
    if not first_col:
        return False

    return parse_csf_ym_from_row(first_col) is not None


def parse_fundgov_csf(pdf_path: Path) -> dict:
    with pdfplumber.open(str(pdf_path)) as pdf:
        if len(pdf.pages) < 2:
            return {}

        page = pdf.pages[1]
        table = page.extract_table()
        if not table:
            return {}

        target_row = None
        for row in reversed(table):
            if is_valid_csf_perf_row(row):
                target_row = row
                break

        if not target_row:
            return {}

        ym = parse_csf_ym_from_row(target_row[0])
        if not ym:
            return {}

        try:
            realized_rate = clean_num(target_row[2])
            overall_rate = clean_num(target_row[4])
        except Exception:
            return {}

        return {
            "ym": ym,
            "退撫基金": realized_rate,
            "退撫基金_整體收益率(%)": overall_rate,
        }


# =========================
# 抓取連結邏輯
# =========================
def collect_blf_links(url):
    links = []
    for p in range(1, 15):
        r = SESSION.get(f"{url}?Page={p}&PageSize=10", timeout=60)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        found = [
            urljoin(url, a["href"])
            for a in soup.find_all("a", href=True)
            if ".pdf" in a["href"].lower() or "media/" in a["href"]
        ]
        if not found:
            break
        links.extend(found)
    return sorted(set(links))


def collect_fundgov_links(url):
    links = []
    for p in range(1, 10):
        r = SESSION.get(f"{url}&page={p}", timeout=60)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        found = [
            urljoin(url, a["href"])
            for a in soup.find_all("a", href=True)
            if "Download.ashx" in a["href"]
        ]
        if not found:
            break
        links.extend(found)
    return sorted(set(links))


# =========================
# 主流程
# =========================
@dataclass
class FundJob:
    name: str
    list_url: str
    collect_fn: callable
    parse_fn: callable
    subdir: str


def main(start_ym: str | None = None):
    if start_ym is not None:
        start_ym = validate_ym_format(start_ym)
        print(f"🚀 開始回補：只處理 {start_ym} 之後（含）的資料")
    else:
        print("🚀 開始回補：處理全部歷史資料")

    jobs = [
        FundJob(
            "勞退",
            "https://www.blf.gov.tw/49200/49255/49261/49269/49273/lpsimplelist",
            collect_blf_links,
            parse_blf_retirement,
            "retirement",
        ),
        FundJob(
            "勞保",
            "https://www.blf.gov.tw/49200/49255/49281/49285/49289/lpsimplelist",
            collect_blf_links,
            parse_blf_labor_ins,
            "labor_ins",
        ),
        FundJob(
            "國保",
            "https://www.blf.gov.tw/49200/49255/49323/49327/49331/lpsimplelist",
            collect_blf_links,
            parse_blf_national_pension,
            "national_pension",
        ),
        FundJob(
            "退撫",
            "https://www.fund.gov.tw/News.aspx?n=658&sms=11737",
            collect_fundgov_links,
            parse_fundgov_csf,
            "csf",
        ),
    ]

    df = pd.DataFrame(columns=COLS)

    for job in jobs:
        print(f"\n📦 開始回補：{job.name}")
        links = job.collect_fn(job.list_url)
        print(f"🔎 找到 {len(links)} 個連結")

        kept = 0
        skipped = 0

        for i, url in enumerate(links, 1):
            try:
                path = download_pdf(url, DOWNLOAD_DIR / job.subdir)
                data = job.parse_fn(path)

                if not data or "ym" not in data:
                    skipped += 1
                    continue

                ym = data["ym"]
                if not should_keep_ym(ym, start_ym):
                    skipped += 1
                    continue

                idx = upsert_row(df, ym)
                for k, v in data.items():
                    if k in COLS:
                        df.loc[idx, k] = v

                kept += 1

                if i % 10 == 0:
                    print(f"  進度: {i}/{len(links)}，保留 {kept} 筆，跳過 {skipped} 筆")

            except Exception as e:
                skipped += 1
                print(f"  ⚠️ 失敗：{url} | {e}")

            time.sleep(0.1)

        print(f"✅ {job.name} 完成：保留 {kept} 筆，跳過 {skipped} 筆")

    df = sort_by_ym(df)
    df = df[COLS]
    df.to_excel(OUT_EXCEL, index=False)
    print(f"\n🎉 回補完成！檔案存於: {OUT_EXCEL}")


if __name__ == "__main__":
    start_ym = sys.argv[1] if len(sys.argv) > 1 else None
    main(start_ym)