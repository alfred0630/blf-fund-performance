# BLF Fund Performance Crawler

📊 **Taiwan public funds performance data crawler**

自動爬取台灣政府公開資料，整理各大公共基金的月度績效，並輸出為結構化 Excel，方便研究、分析與長期追蹤。

目前支援基金：
- 勞工退休基金（新制 / 舊制）
- 勞工保險基金
- 國民年金保險基金（含比較指標）
- 退撫基金（整體實際收益率）

---

## 🧠 專案功能

### ✅ 每月自動更新
- 下載各基金最新公布的 PDF
- 解析指定績效欄位
- 合併後 **只寫入一次 Excel**（避免鎖檔與資料覆蓋問題）

### 🧰 歷史資料回補（Backfill）
- 批次抓取各站所有歷史 PDF
- 逐月解析並合併
- 輸出為 **獨立的歷史 Excel 檔**

---

## 📊 輸出資料欄位

| 欄位名稱 | 說明 |
|--------|------|
| 年月 | 資料所屬年月 |
| 新制基金_YTD(%) | 勞退新制 年初至今報酬率 |
| 舊制基金_YTD(%) | 勞退舊制 年初至今報酬率 |
| 勞保基金_YTD(%) | 勞保基金 年初至今報酬率 |
| 國保基金_YTD(%) | 國民年金保險基金 年初至今報酬率 |
| TW10Y_YTD(%) | 台灣 10 年期公債殖利率（YTD） |
| TW_TAIEX_YTD(%) | 台灣加權股價指數（YTD） |
| MSCI_World_YTD(%) | MSCI 全球股票指數（YTD） |
| BBG_GlobalBond_YTD(%) | Bloomberg 全球債券指數（YTD） |
| 退撫基金_整體實際收益率(%) | 退撫基金整體實際收益率 |

---

## 🚀 安裝與環境

### 1️⃣ Clone 專案

```bash
git clone https://github.com/alfred0630/blf-fund-performance.git
cd blf-fund-performance
```
### 2️⃣ 建立虛擬環境（建議）

```bash
python -m venv .venv
```
# Windows
.\.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

### 3️⃣ 安裝套件
```bash

pip install -r requirements.txt
requirements.txt

nginx
requests
beautifulsoup4
pdfplumber
pandas
openpyxl
```
### 📁 專案結構
```bash

blf-fund-performance/
├─ data_pdf/                     # PDF 下載資料
├─ blf_returns.xlsx              # 每月更新輸出
├─ blf_returns_backfill.xlsx     # 歷史回補輸出
├─ common.py                     # 共用工具（路徑 / 年月處理）
├─ excel_store.py                # Excel 單次寫入模組
├─ blf_download_retirement.py
├─ blf_parse_retirement.py
├─ blf_download_labor_ins.py
├─ blf_parse_labor_ins.py
├─ blf_download_national_pension.py
├─ blf_parse_national_pension.py
├─ fundgov_download_csf.py
├─ fundgov_parse_csf.py
├─ run_all_pro.py                # ✅ 每月自動更新（專業版）
└─ backfill_history.py           # 🧰 一次性歷史回補

```
### ▶️ 使用方式
#### 🆕 每月更新最新資料
```bash
python run_all_pro.py
```
#### 流程：

下載各基金最新 PDF

解析績效數據

合併寫入 blf_returns.xlsx

🧰 回補歷史資料（只需跑一次）
```bash
python backfill_history.py
```

輸出：
```
blf_returns_backfill.xlsx
```
📌 不會影響每月更新用的 Excel。

⚙️ 可攜性與自動化
全部使用 相對路徑

可直接複製到其他電腦或伺服器

可搭配排程工具使用

Windows 工作排程（範例）
```powershell

schtasks /create /sc monthly /mo 1 /tn "BLF Fund Update" /tr "python C:\path\to\run_all_pro.py"
```
