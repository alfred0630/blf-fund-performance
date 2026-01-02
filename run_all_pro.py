# run_all_pro.py  (professional: download -> parse -> merge in-memory -> write Excel once)

from common import BASE_DIR
from excel_store import load_excel, ensure_columns, upsert_row, sort_by_ym, safe_write_excel

from blf_download_retirement import download_latest as dl_ret
from blf_parse_retirement import extract as parse_ret

from blf_download_labor_ins import download_latest as dl_li
from blf_parse_labor_ins import extract as parse_li

from blf_download_national_pension import download_latest as dl_npi
from blf_parse_national_pension import extract as parse_npi

from fundgov_download_csf import download_latest as dl_csf
from fundgov_parse_csf import extract as parse_csf


EXCEL_PATH = BASE_DIR / "blf_returns.xlsx"

# ✅ 欄位統一：語意清楚、避免誤會
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

def require_ym(payload: dict, name: str):
    ym = payload.get("ym")
    if not ym:
        raise ValueError(f"{name} 解析到的 ym 是空的，請檢查 PDF 格式或解析規則。")
    return ym

def apply_retirement(df, payload):
    ym = require_ym(payload, "勞退")
    idx = upsert_row(df, ym)
    df.loc[idx, "新制基金_YTD(%)"] = payload["retirement_new_ytd"]
    df.loc[idx, "舊制基金_YTD(%)"] = payload["retirement_old_ytd"]

def apply_labor_ins(df, payload):
    ym = require_ym(payload, "勞保")
    idx = upsert_row(df, ym)
    df.loc[idx, "勞保基金_YTD(%)"] = payload["labor_ins_ytd"]

def apply_national_pension(df, payload):
    ym = require_ym(payload, "國保")
    idx = upsert_row(df, ym)
    df.loc[idx, "國保基金_YTD(%)"] = payload["npi_ytd"]
    df.loc[idx, "TW10Y_YTD(%)"] = payload["tw10y_ytd"]
    df.loc[idx, "TW_TAIEX_YTD(%)"] = payload["taiex_ytd"]
    df.loc[idx, "MSCI_World_YTD(%)"] = payload["msci_world_ytd"]
    df.loc[idx, "BBG_GlobalBond_YTD(%)"] = payload["bbg_globalbond_ytd"]

def apply_csf(df, payload):
    ym = require_ym(payload, "退撫")
    idx = upsert_row(df, ym)
    # ✅ 退撫基金抓的是「整體實際收益率(%)」
    df.loc[idx, "退撫基金_整體實際收益率(%)"] = payload["csf_overall_actual_return"]

def main():
    print("🚀 專業版：下載 → 解析 → 記憶體合併 → 最後一次寫入 Excel")

    # 1) 下載
    ret_pdf = dl_ret()
    li_pdf = dl_li()
    npi_pdf = dl_npi()
    csf_pdf = dl_csf()

    # 2) 解析（回傳 dict）
    ret = parse_ret(ret_pdf)
    li = parse_li(li_pdf)
    npi = parse_npi(npi_pdf)
    csf = parse_csf(csf_pdf)

    # 3) 載入 Excel + 合併更新（只在記憶體，不寫檔）
    df = load_excel(EXCEL_PATH)
    df = ensure_columns(df, COLS)

    # 同月 upsert：永遠覆蓋更新（不跳過）
    apply_retirement(df, ret)
    apply_labor_ins(df, li)
    apply_national_pension(df, npi)
    apply_csf(df, csf)

    # 4) 排序 + 一次寫入
    df = sort_by_ym(df)
    out_path = safe_write_excel(df, EXCEL_PATH)

    print(f"🎉 完成！寫入：{out_path}")

if __name__ == "__main__":
    main()
