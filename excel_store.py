from pathlib import Path
import pandas as pd

SHEET_NAME = "data"

def upsert_row(df: pd.DataFrame, ym: str) -> int:
    if "年月" not in df.columns:
        df["年月"] = pd.Series(dtype="object")

    if (df["年月"] == ym).any():
        return df.index[df["年月"] == ym][0]

    idx = len(df)
    df.loc[idx, "年月"] = ym
    return idx

def ensure_columns(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def load_excel(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_excel(path, sheet_name=SHEET_NAME)
    return pd.DataFrame(columns=["年月"])

def sort_by_ym(df: pd.DataFrame) -> pd.DataFrame:
    df["年月"] = pd.to_datetime(df["年月"])
    df = df.sort_values("年月").reset_index(drop=True)
    df["年月"] = df["年月"].dt.strftime("%Y-%m")
    return df

def safe_write_excel(df: pd.DataFrame, path: Path):
    """
    寫入時如果 Excel 正在開，會 PermissionError。
    專業版做法：鎖住就另存一份帶時間戳的檔案，流程不中斷。
    """
    out_path = path
    try:
        # 測試是否可寫（Windows 常見鎖檔）
        with open(path, "a"):
            pass
    except PermissionError:
        ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        out_path = path.with_name(f"{path.stem}_{ts}{path.suffix}")
        print(f"⚠️ Excel 可能正在開著，改寫到：{out_path.name}")

    with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    return out_path
