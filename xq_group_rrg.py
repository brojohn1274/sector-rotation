#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xq_group_rrg.py
================
把 XQ 全球贏家「族群透視」細產業/主題族群的資料，轉成跟 twse_sector_rrg.html
相容的 RRG (RS-Ratio / RS-Momentum) JSON。

因為 XQ 沒有開放 API，這支程式改吃「你從 XQ 手動匯出的檔案」，支援兩種情境：

情境 A（建議）：每個細產業有自己的歷史指數走勢
--------------------------------------------
XQ 的「細產業指標(系統)」本身是一個可以開技術分析線圖的商品，
在該商品的 K 線視窗按左鍵 →「輸出到 Excel」，會匯出一份含日期/收盤價的歷史檔。
把每個族群匯出的檔案都丟進同一個資料夾（一個族群一個檔案，檔名= 族群名稱），
例如：
    xq_export/
        AI伺服器.xlsx
        PCB.xlsx
        CPO.xlsx
        半導體.xlsx
        ...
執行：
    python xq_group_rrg.py groups --input-dir xq_export --json-out xq_group_rrg.json

情境 B（備用，也是目前用的模式）：只有成分股清單，沒有現成指數
--------------------------------------------
如果族群透視只能匯出「族群 -> 成分股」清單，沒有現成指數走勢，
就用等權重平均漲跌幅自己合成一個族群指數（不需要市值資料，較簡化）。
只要準備一份 mapping CSV，兩欄：group,stock_name（直接用中文股票名稱即可，
不需要代號）：

    group,stock_name
    半導體設備指標,弘塑
    半導體設備指標,家登
    IC封測指標,日月光投控
    ...

程式會自動下載 TWSE 上市 + TPEx 上櫃兩份官方公司清單（代號/名稱對照表），
把每個股票名稱比對成代號跟市場別，再抓歷史股價（上市走 TWSE 端點、
上櫃走 TPEx 端點，兩邊格式不同，程式會自動判斷）。名稱比對用「完全一致」
優先，找不到會列出來，你可以手動補正（例如全形/半形字元不同、簡稱不同、
興櫃/公開發行股不在上市櫃清單裡本來就會對不到）。

執行（會自動用 TWSE 免費個股日成交資料回補股價）：
    python xq_group_rrg.py constituents --mapping xq_groups.csv \\
        --start 2026-05-01 --end 2026-08-01 \\
        --json-out xq_group_rrg.json

兩種模式最後都會產生同一種 JSON 格式，直接用 twse_sector_rrg.html
的「上傳 JSON」按鈕載入即可（不需要改 HTML）。

基準指數（大盤）
----------------
兩種模式都會自動用 TWSE 免費端點抓「發行量加權股價指數」當基準
（不需要你額外從 XQ 匯出大盤資料），除非你用 --benchmark-file 指定
自己匯出的大盤檔案。

同樣提醒：這支程式需要在你自己的機器上執行，我這邊的沙盒對外網域
是被限制的，無法在這裡直接幫你測試 XQ 匯出檔或連線 TWSE，
邏輯已依常見匯出格式撰寫，第一次用建議先拿 2-3 個族群小量測試。
"""

import argparse
import glob
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta

import pandas as pd
import requests

TWSE_STOCK_DAY = "https://www.twse.com.tw/exchangeReport/STOCK_DAY"
TWSE_MI_INDEX = "https://www.twse.com.tw/exchangeReport/MI_INDEX"
TWSE_ISIN_LIST = "https://isin.twse.com.tw/isin/C_public.jsp?strMode=2"  # 上市公司代號/名稱對照表
TPEX_ISIN_LIST = "https://isin.twse.com.tw/isin/C_public.jsp?strMode=4"  # 上櫃公司代號/名稱對照表
TPEX_DAILY_QUOTE = "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php"  # 上櫃全市場單日收盤


def _fetch_isin_list(url: str, market_label: str) -> dict:
    """下載單一 strMode 的公司代號/名稱對照表（TWSE 上市或 TPEx 上櫃通用）。
    回傳 {股票名稱: {"code":代號, "industry":官方產業別, "market":"TWSE"/"TPEx"}}。

    這個頁面是很舊式、不太合乎規範的 HTML（屬性值沒加引號等），pandas.read_html
    在解析時常會失敗、甚至把解析錯誤訊息裡帶著整段原始 HTML 印出來。
    改用正則表達式直接從原始 HTML 文字裡挖出 <tr>/<td>，對這種不規範格式更耐操。
    編碼用 cp950（是 big5 的超集，含少數 big5 沒有的字，公司名稱裡偶爾會用到）。
    """
    print(f"下載 {market_label} 公司代號/名稱/產業別對照表 ...")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.encoding = "cp950"
        html = resp.text
    except Exception as e:
        raise RuntimeError(f"無法下載 {market_label} 公司清單: {e}")

    def strip_tags(s: str) -> str:
        s = re.sub(r"<[^>]*>", "", s)
        s = s.replace("&nbsp;", " ").replace("&amp;", "&")
        return s.strip()

    name_to_info = {}
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I)
    for row_html in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row_html, re.S | re.I)
        if len(cells) < 5:
            continue
        cells = [strip_tags(c) for c in cells]
        first_col = cells[0]
        # 代號和名稱中間通常夾了全形空白，偶爾是半形空白
        parts = first_col.split("\u3000")
        if len(parts) < 2:
            parts = first_col.split(" ", 1)
        if len(parts) < 2:
            continue
        code, name = parts[0].strip(), parts[1].strip()
        industry = cells[4].strip() if len(cells) > 4 else ""
        # 產業別欄位空白代表這不是普通個股（權證、受益證券、TDR、部分債券等），
        # 這些不該被當成個股去抓歷史股價，直接跳過，不要塞進「其他」分類。
        if code.isdigit() and name and industry:
            name_to_info[name] = {
                "code": code,
                "industry": industry,
                "market": market_label,
            }

    if not name_to_info:
        raise RuntimeError(
            f"{market_label} 公司清單解析結果是空的，可能是網站格式又變了，"
            "把這個錯誤訊息回報給我，我再調整解析邏輯。"
        )
    print(f"  取得 {len(name_to_info)} 檔{market_label}公司代號/產業別對照")
    return name_to_info


def build_name_to_code_map() -> dict:
    """合併 TWSE 上市 + TPEx 上櫃兩份公司清單。
    回傳 {股票名稱: {"code":代號, "industry":官方產業別, "market":"TWSE"/"TPEx"}}。
    上市與上櫃股票名稱理論上不會重複，若真的撞名以上市優先。
    """
    twse_map = _fetch_isin_list(TWSE_ISIN_LIST, "TWSE上市")
    tpex_map = _fetch_isin_list(TPEX_ISIN_LIST, "TPEx上櫃")
    merged = dict(tpex_map)
    merged.update(twse_map)  # 上市優先覆蓋（如果撞名）
    print(f"  合併後共 {len(merged)} 檔（上市 {len(twse_map)} + 上櫃 {len(tpex_map)}）")
    return merged


def resolve_names_to_codes(names: list) -> dict:
    """把股票名稱清單解析成 {名稱: {"code":代號, "market":市場別}}，找不到的會列出來提醒使用者。"""
    name_to_info = build_name_to_code_map()
    resolved, missing = {}, []
    for n in names:
        n_clean = n.strip()
        info = name_to_info.get(n_clean) or name_to_info.get(n_clean.rstrip("*").strip())
        if info:
            resolved[n_clean] = {"code": info["code"], "market": info["market"]}
        else:
            missing.append(n_clean)
    if missing:
        print(f"[warn] 有 {len(missing)} 檔股票名稱對不到代號，請手動確認: {missing}")
    return resolved


def build_full_market_mapping(fine_override_csv: str, out_csv: str = "xq_groups_full.csv"):
    """產生全市場 group,stock_name 對照表：
    - 有在 fine_override_csv 裡的股票 -> 用我整理的細分類（AI伺服器_散熱 這種）
    - 其他全部股票 -> 用 TWSE 官方「產業別」（傳產/金融自動涵蓋，全市場約1700檔）
    """
    name_to_info = build_name_to_code_map()

    override = pd.read_csv(fine_override_csv, dtype=str)
    override.columns = [c.strip().lower() for c in override.columns]
    override_map = dict(zip(override["stock_name"].str.strip(), override["group"].str.strip()))

    rows = []
    for name, info in name_to_info.items():
        group = override_map.get(name) or override_map.get(name.rstrip("*")) or info["industry"]
        rows.append({"group": group, "stock_name": name})

    out = pd.DataFrame(rows)
    out.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"已產生全市場分類表 -> {out_csv}（{out['group'].nunique()} 個分類，{len(out)} 檔股票）")
    return out
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

DATE_CANDIDATES = ["日期", "成交日期", "date", "Date"]
CLOSE_CANDIDATES = ["收盤", "收盤價", "收盤指數", "close", "Close", "收盤(元)"]


# ---------------------------------------------------------------------
# 情境 A：讀取每個族群自己的歷史匯出檔
# ---------------------------------------------------------------------

def _read_any(path: str) -> pd.DataFrame:
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)
    # XQ 匯出的 CSV 常見是 Big5 或 UTF-8-SIG，兩種都試
    for enc in ("utf-8-sig", "big5", "cp950", "utf-8"):
        try:
            return pd.read_csv(path, encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise RuntimeError(f"無法讀取檔案編碼: {path}")


def _find_col(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    # 模糊比對：欄名包含關鍵字
    for c in df.columns:
        for cand in candidates:
            if cand in str(c):
                return c
    return None


def load_group_file(path: str) -> pd.Series:
    """讀一個族群的歷史匯出檔，回傳 date-indexed 收盤值 Series"""
    df = _read_any(path)
    date_col = _find_col(df, DATE_CANDIDATES)
    close_col = _find_col(df, CLOSE_CANDIDATES)
    if date_col is None or close_col is None:
        raise RuntimeError(
            f"{path} 找不到日期/收盤欄位，實際欄位有: {list(df.columns)}\n"
            f"請確認匯出格式，或手動改欄名成「日期」「收盤」再試一次。"
        )
    dates = pd.to_datetime(df[date_col].astype(str).str.replace("/", "-"), errors="coerce")
    closes = pd.to_numeric(
        df[close_col].astype(str).str.replace(",", "").str.replace("=", ""), errors="coerce"
    )
    s = pd.Series(closes.values, index=dates.values).dropna()
    s = s[~s.index.isna()]
    s.index = pd.to_datetime(s.index)
    return s.sort_index()


def build_from_group_files(input_dir: str) -> pd.DataFrame:
    files = sorted(
        glob.glob(os.path.join(input_dir, "*.xlsx"))
        + glob.glob(os.path.join(input_dir, "*.xls"))
        + glob.glob(os.path.join(input_dir, "*.csv"))
    )
    if not files:
        raise RuntimeError(f"{input_dir} 底下沒有找到任何 .xlsx/.xls/.csv 檔案")

    series_map = {}
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        print(f"讀取族群「{name}」 <- {f}")
        try:
            series_map[name] = load_group_file(f)
        except Exception as e:
            print(f"  [warn] 略過 {f}: {e}")

    df = pd.DataFrame(series_map)
    df.index.name = "date"
    return df.sort_index()


# ---------------------------------------------------------------------
# 情境 B：只有成分股清單 -> 等權重合成族群指數
# ---------------------------------------------------------------------

def fetch_twse_day_all(d: datetime, retries=3, pause=1.0) -> dict:
    """抓單一交易日「全部上市股票」的收盤價（TWSE MI_INDEX ALLBUT0999），回傳 {代號: 收盤價}"""
    date_str = d.strftime("%Y%m%d")
    params = {"response": "json", "date": date_str, "type": "ALLBUT0999"}
    for attempt in range(retries):
        try:
            r = requests.get(TWSE_MI_INDEX, params=params, headers=HEADERS, timeout=10)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"  [warn] TWSE {date_str} 第 {attempt+1} 次失敗: {e}")
            time.sleep(pause)
            continue
        if data.get("stat") not in ("OK", "ok"):
            return {}
        # 「每日收盤行情」表通常叫 data9，欄位常見為
        # [證券代號, 證券名稱, 成交股數, 成交筆數, 成交金額, 開盤價, 最高價, 最低價, 收盤價, ...]
        rows = data.get("data9")
        if rows is None:
            # 版型偶爾會變，退而求其次去 tables 裡找欄位含「收盤價」的那張表
            for t in data.get("tables", []):
                fields = t.get("fields", [])
                if any("收盤價" in f for f in fields):
                    rows = t.get("data", [])
                    break
        if not rows:
            return {}
        out = {}
        for row in rows:
            try:
                code = str(row[0]).strip()
                close = float(str(row[8]).replace(",", ""))
                out[code] = close
            except (ValueError, IndexError, TypeError):
                continue
        return out
    return {}


def fetch_twse_history_batch(codes: set, start: datetime, end: datetime, sleep_sec=0.5) -> dict:
    """逐日抓 TWSE 全市場收盤價，只保留需要的代號。回傳 {代號: pd.Series}"""
    per_code = {c: {} for c in codes}
    d = start
    n_days = 0
    while d <= end:
        if d.weekday() < 5:
            day_data = fetch_twse_day_all(d)
            n_days += 1
            date_str = d.strftime("%Y-%m-%d")
            for c in codes:
                if c in day_data:
                    per_code[c][date_str] = day_data[c]
            time.sleep(sleep_sec)
        d += timedelta(days=1)
    print(f"  TWSE 上市：共查詢 {n_days} 個交易日")
    result = {}
    for c, data in per_code.items():
        if data:
            s = pd.Series(data)
            s.index = pd.to_datetime(s.index)
            result[c] = s.sort_index()
    return result


def fetch_stock_month(stock_no: str, year: int, month: int, retries=3, pause=1.0):
    """抓單一股票、單一月份的日收盤（TWSE STOCK_DAY，免費，一次回傳整月）
    保留這個函式供 groups 模式或個別除錯使用；constituents 模式已改用
    fetch_twse_history_batch（單日全市場批次）以大幅減少請求數。"""
    date_str = f"{year}{month:02d}01"
    params = {"response": "json", "date": date_str, "stockNo": stock_no}
    for attempt in range(retries):
        try:
            r = requests.get(TWSE_STOCK_DAY, params=params, headers=HEADERS, timeout=10)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"  [warn] {stock_no} {year}-{month:02d} 第 {attempt+1} 次失敗: {e}")
            time.sleep(pause)
            continue
        if data.get("stat") not in ("OK", "ok") or "data" not in data:
            return {}
        fields = data.get("fields", [])
        try:
            d_idx = fields.index("日期")
            c_idx = fields.index("收盤價")
        except ValueError:
            d_idx, c_idx = 0, 6
        out = {}
        for row in data["data"]:
            try:
                roc_date = row[d_idx]  # 民國年日期 like "115/06/03"
                y, m, d = roc_date.split("/")
                greg_date = f"{int(y)+1911:04d}-{int(m):02d}-{int(d):02d}"
                close = float(str(row[c_idx]).replace(",", ""))
                out[greg_date] = close
            except (ValueError, IndexError):
                continue
        return out
    return {}


def fetch_stock_history(stock_no: str, start: datetime, end: datetime, sleep_sec=0.5) -> pd.Series:
    months = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    data = {}
    for (y, m) in months:
        data.update(fetch_stock_month(stock_no, y, m))
        time.sleep(sleep_sec)
    s = pd.Series(data)
    s.index = pd.to_datetime(s.index)
    return s.sort_index()


# ---------------------------------------------------------------------
# TPEx（上櫃）股價：跟 TWSE 不同，官方端點是「單日回傳全市場」，
# 不是「單股回傳整月」，所以用逐日批次抓、一次取出所有需要的代號。
# ---------------------------------------------------------------------

def fetch_tpex_day_all(d: datetime, retries=3, pause=1.0) -> dict:
    """抓單一交易日「全部上櫃股票」的收盤價，回傳 {代號: 收盤價}"""
    roc = f"{d.year - 1911}/{d.month:02d}/{d.day:02d}"
    params = {"l": "zh-tw", "d": roc}
    for attempt in range(retries):
        try:
            r = requests.get(TPEX_DAILY_QUOTE, params=params, headers=HEADERS, timeout=10)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"  [warn] TPEx {roc} 第 {attempt+1} 次失敗: {e}")
            time.sleep(pause)
            continue
        rows = data.get("aaData") or data.get("tables", [{}])[0].get("data", []) if isinstance(data, dict) else []
        out = {}
        for row in rows:
            try:
                code = str(row[0]).strip()
                close = float(str(row[2]).replace(",", ""))
                out[code] = close
            except (ValueError, IndexError, TypeError):
                continue
        return out
    return {}


def fetch_tpex_history_batch(codes: set, start: datetime, end: datetime, sleep_sec=0.5) -> dict:
    """逐日抓 TPEx 全市場收盤價，只保留需要的代號。回傳 {代號: pd.Series}"""
    per_code = {c: {} for c in codes}
    d = start
    n_days = 0
    while d <= end:
        if d.weekday() < 5:
            day_data = fetch_tpex_day_all(d)
            n_days += 1
            date_str = d.strftime("%Y-%m-%d")
            for c in codes:
                if c in day_data:
                    per_code[c][date_str] = day_data[c]
            time.sleep(sleep_sec)
        d += timedelta(days=1)
    print(f"  TPEx 上櫃：共查詢 {n_days} 個交易日")
    result = {}
    for c, data in per_code.items():
        if data:
            s = pd.Series(data)
            s.index = pd.to_datetime(s.index)
            result[c] = s.sort_index()
    return result


def build_from_constituents(mapping_csv: str, start: str, end: str):
    mapping = pd.read_csv(mapping_csv, dtype=str)
    mapping.columns = [c.strip().lower() for c in mapping.columns]
    if "group" not in mapping.columns:
        raise RuntimeError("mapping CSV 需要有 group 欄")
    if "stock_code" not in mapping.columns and "stock_name" not in mapping.columns:
        raise RuntimeError("mapping CSV 需要有 stock_code 或 stock_name 其中一欄")

    if "stock_code" not in mapping.columns:
        # 只有名稱 -> 自動解析代號 + 市場別（TWSE上市 / TPEx上櫃）
        all_names = sorted(mapping["stock_name"].unique())
        resolved = resolve_names_to_codes(all_names)
        mapping["stock_code"] = mapping["stock_name"].map(lambda n: resolved.get(n, {}).get("code"))
        mapping["market"] = mapping["stock_name"].map(lambda n: resolved.get(n, {}).get("market"))
        before = len(mapping)
        mapping = mapping.dropna(subset=["stock_code"])
        dropped = before - len(mapping)
        if dropped:
            print(f"[warn] 有 {dropped} 筆因為對不到代號被跳過")
    elif "market" not in mapping.columns:
        # 已經有 stock_code 但沒標市場別，預設當 TWSE 上市處理（向下相容舊格式）
        mapping["market"] = "TWSE上市"

    if "stock_name" not in mapping.columns:
        mapping["stock_name"] = mapping["stock_code"]

    # 保留每個族群的成分股名稱清單，供 HTML 表格模式的「成份」彈窗使用
    constituents = {g: sorted(sub["stock_name"].unique().tolist()) for g, sub in mapping.groupby("group")}

    start_dt = datetime.strptime(start, "%Y-%m-%d")
    end_dt = datetime.strptime(end, "%Y-%m-%d")

    twse_codes = sorted(mapping.loc[mapping["market"] != "TPEx上櫃", "stock_code"].unique())
    tpex_codes = sorted(mapping.loc[mapping["market"] == "TPEx上櫃", "stock_code"].unique())

    price_cache = {}
    if twse_codes:
        print(f"抓取上市個股歷史股價（共 {len(twse_codes)} 檔，逐日批次抓取）...")
        price_cache.update(fetch_twse_history_batch(set(twse_codes), start_dt, end_dt))

    if tpex_codes:
        print(f"抓取上櫃個股歷史股價（共 {len(tpex_codes)} 檔，逐日批次抓取）...")
        price_cache.update(fetch_tpex_history_batch(set(tpex_codes), start_dt, end_dt))

    group_series = {}
    for group, sub in mapping.groupby("group"):
        codes = list(sub["stock_code"])
        # 等權重：每檔股票各自漲跌幅取平均，再轉成指數化(基期=100)
        rets = []
        for code in codes:
            s = price_cache.get(code)
            if s is None or s.empty:
                continue
            r = s / s.iloc[0] * 100
            rets.append(r.rename(code))
        if not rets:
            continue
        combined = pd.concat(rets, axis=1)
        group_series[group] = combined.mean(axis=1)

    df = pd.DataFrame(group_series)
    df.index.name = "date"
    return df.sort_index(), constituents


# ---------------------------------------------------------------------
# 基準指數（大盤）：預設自動抓 TWSE 加權指數；也可用檔案指定
# ---------------------------------------------------------------------

def fetch_benchmark_auto(start: datetime, end: datetime, sleep_sec=0.8) -> pd.Series:
    out = {}
    d = start
    while d <= end:
        if d.weekday() < 5:
            date_str = d.strftime("%Y%m%d")
            params = {"response": "json", "date": date_str, "type": "IND"}
            try:
                r = requests.get(TWSE_MI_INDEX, params=params, headers=HEADERS, timeout=10)
                r.raise_for_status()
                data = r.json()
                if data.get("stat") in ("OK", "ok"):
                    rows, fields = None, None
                    if "tables" in data:
                        for t in data["tables"]:
                            f = t.get("fields", [])
                            if any("收盤指數" in x for x in f):
                                rows, fields = t.get("data", []), f
                                break
                    if rows:
                        name_idx = next((i for i, f in enumerate(fields) if "指數名稱" in f), 0)
                        close_idx = next((i for i, f in enumerate(fields) if "收盤指數" in f), 1)
                        for row in rows:
                            if "加權" in str(row[name_idx]):
                                out[d.strftime("%Y-%m-%d")] = float(str(row[close_idx]).replace(",", ""))
                                break
            except Exception as e:
                print(f"  [warn] 基準指數 {date_str} 抓取失敗: {e}")
            time.sleep(sleep_sec)
        d += timedelta(days=1)
    s = pd.Series(out)
    s.index = pd.to_datetime(s.index)
    return s.sort_index()


# ---------------------------------------------------------------------
# RRG 計算（跟 twse_sector_rrg.py 相同公式）
# ---------------------------------------------------------------------

def zscore(series: pd.Series, window: int) -> pd.Series:
    roll_mean = series.rolling(window).mean()
    roll_std = series.rolling(window).std(ddof=0)
    return (series - roll_mean) / roll_std.replace(0, pd.NA)


def compute_rrg(df: pd.DataFrame, bench: pd.Series, smooth=5, window=10, momentum_period=5, constituents=None) -> dict:
    df = df.copy()
    bench = bench.reindex(df.index).ffill()
    result = {"dates": [d.strftime("%Y-%m-%d") for d in df.index], "sectors": {}}
    if constituents:
        result["constituents"] = constituents

    for col in df.columns:
        rs = 100 * df[col] / bench
        rs_smooth = rs.ewm(span=smooth, adjust=False).mean()
        rs_ratio = 100 + zscore(rs_smooth, window)
        roc = rs_ratio.diff(momentum_period)
        rs_momentum = 100 + zscore(roc, window)

        points = []
        for dt, ratio, mom in zip(df.index, rs_ratio, rs_momentum):
            if pd.isna(ratio) or pd.isna(mom):
                continue
            points.append({"date": dt.strftime("%Y-%m-%d"), "ratio": round(float(ratio), 3), "momentum": round(float(mom), 3)})
        if points:
            result["sectors"][col] = points
    return result


# 計算週期預設（對應網頁上的 20短線/60波段/120中期/240長期按鈕）
# smooth=RS平滑天數, window=zscore滾動視窗, momentum_period=動能變化率天數
PERIOD_PRESETS = {
    "20": {"label": "20 短線", "smooth": 3, "window": 10, "momentum_period": 3},
    "60": {"label": "60 波段", "smooth": 5, "window": 20, "momentum_period": 5},
    "120": {"label": "120 中期", "smooth": 10, "window": 40, "momentum_period": 10},
    "240": {"label": "240 長期", "smooth": 15, "window": 60, "momentum_period": 15},
}


def compute_all_variants(df: pd.DataFrame, constituents=None) -> dict:
    """算出「加權指數」「等權類股」兩種基準 x 20/60/120/240 四種週期，共 8 組，
    輸出成 twse_sector_rrg.html 看得懂的巢狀格式：
        {"variants": {benchmark_key: {period_key: {dates, sectors, constituents}}},
         "benchmarks": {...labels...}, "periods": {...labels...}}
    """
    print("自動抓取 TWSE 大盤(發行量加權股價指數)當基準 ...")
    bench_weighted = fetch_benchmark_auto(df.index.min(), df.index.max())
    bench_equal = df.mean(axis=1)  # 等權類股：所有族群指數的簡單平均，當作另一種基準

    benchmarks = {"weighted": bench_weighted, "equal": bench_equal}
    benchmark_labels = {"weighted": "加權指數", "equal": "等權類股"}

    variants = {}
    for bkey, bseries in benchmarks.items():
        variants[bkey] = {}
        for pkey, preset in PERIOD_PRESETS.items():
            print(f"計算 基準={benchmark_labels[bkey]} 週期={preset['label']} ...")
            variants[bkey][pkey] = compute_rrg(
                df, bseries,
                smooth=preset["smooth"], window=preset["window"], momentum_period=preset["momentum_period"],
                constituents=constituents,
            )

    return {
        "variants": variants,
        "benchmarks": benchmark_labels,
        "periods": {k: v["label"] for k, v in PERIOD_PRESETS.items()},
        "default_benchmark": "weighted",
        "default_period": "60",
    }


# ---------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="XQ 族群透視 -> RRG 產業輪動圖 資料產生器")
    sub = ap.add_subparsers(dest="mode", required=True)

    p_groups = sub.add_parser("groups", help="情境A：每族群有自己的歷史匯出檔")
    p_groups.add_argument("--input-dir", required=True)
    p_groups.add_argument("--benchmark-file", default=None, help="自己匯出的大盤歷史檔（不指定則自動抓TWSE加權指數）")
    p_groups.add_argument("--json-out", default="xq_group_rrg.json")

    p_const = sub.add_parser("constituents", help="情境B：只有成分股清單，等權重合成")
    p_const.add_argument("--mapping", required=True)
    p_const.add_argument("--start", required=True)
    p_const.add_argument("--end", required=True)
    p_const.add_argument("--json-out", default="xq_group_rrg.json")

    p_full = sub.add_parser("full-market", help="產生全市場分類表（細分類覆蓋 + TWSE官方產業別補齊）")
    p_full.add_argument("--fine-override", default="xq_groups.csv", help="我整理的細分類 CSV（group,stock_name）")
    p_full.add_argument("--out", default="xq_groups_full.csv")

    args = ap.parse_args()

    if args.mode == "groups":
        df = build_from_group_files(args.input_dir)
        constituents = None
    elif args.mode == "full-market":
        build_full_market_mapping(args.fine_override, args.out)
        return
    else:
        df, constituents = build_from_constituents(args.mapping, args.start, args.end)

    if df.empty:
        raise RuntimeError("沒有組出任何族群資料，請檢查輸入檔案")

    # 現在一次算出「加權指數/等權類股」x「20/60/120/240日」共 8 組，讓網頁可以即時切換
    rrg = compute_all_variants(df, constituents=constituents)
    with open(args.json_out, "w", encoding="utf-8") as f:
        json.dump(rrg, f, ensure_ascii=False)
    n_sectors = len(next(iter(next(iter(rrg["variants"].values())).values()))["sectors"])
    print(f"已輸出 {args.json_out}（族群數={n_sectors}，共 {len(rrg['variants'])}種基準 x {len(PERIOD_PRESETS)}種週期）")
    print("接著在 twse_sector_rrg.html 按「上傳 JSON」載入即可，網頁上可以直接切換基準/週期/尾巴長度。")


if __name__ == "__main__":
    sys.exit(main())
