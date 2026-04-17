#!/usr/bin/env python3
"""
每天自動從 TWSE 抓公開申購資料，輸出 data.json 供網頁使用。
改用 TWSE 的 JSON 格式 API，更穩定。
"""
import json
import urllib.request
import urllib.parse
from datetime import datetime, date

def fetch_url(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Accept": "application/json, text/html, */*",
        "Referer": "https://www.twse.com.tw/",
    })
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8", errors="replace")

current_roc_year = date.today().year - 1911
years_to_fetch = [current_roc_year - 1, current_roc_year]

all_rows = []

for yy in years_to_fetch:
    # 用 JSON 格式抓（比解析 HTML 穩定）
    url = f"https://www.twse.com.tw/announcement/publicForm?response=json&yy={yy}"
    try:
        raw = fetch_url(url)
        data = json.loads(raw)
        rows = data.get("data", [])
        print(f"[INFO] {yy} 年 JSON 格式：{len(rows)} 筆")
        all_rows.extend(rows)
        continue
    except Exception as e:
        print(f"[WARN] {yy} 年 JSON 格式失敗: {e}，改用 HTML 解析")

    # fallback：HTML 格式
    try:
        import re
        url2 = f"https://www.twse.com.tw/announcement/publicForm?response=html&yy={yy}"
        html = fetch_url(url2)
        rows2 = re.findall(r"<tr>(.*?)</tr>", html, re.DOTALL)
        count = 0
        for row in rows2:
            cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
            cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            if len(cells) >= 17:
                try:
                    int(cells[0])
                    all_rows.append(cells)
                    count += 1
                except ValueError:
                    pass
        print(f"[INFO] {yy} 年 HTML 格式：{count} 筆")
    except Exception as e2:
        print(f"[ERROR] {yy} 年完全失敗: {e2}")

print(f"[INFO] 共抓到 {len(all_rows)} 筆原始資料")

# ── 如果兩種格式都失敗，試備用 URL ──
if len(all_rows) == 0:
    print("[WARN] 主要 URL 失敗，嘗試備用方式...")
    try:
        import re
        for yy in years_to_fetch:
            url3 = f"https://www.twse.com.tw/rwd/zh/announcement/public?type=public&response=html&yy={yy}"
            try:
                html = fetch_url(url3)
                rows3 = re.findall(r"<tr>(.*?)</tr>", html, re.DOTALL)
                for row in rows3:
                    cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
                    cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                    if len(cells) >= 17:
                        try:
                            int(cells[0])
                            all_rows.append(cells)
                        except ValueError:
                            pass
                print(f"[INFO] 備用 {yy} 年：{len(all_rows)} 筆")
            except Exception as e3:
                print(f"[WARN] 備用 {yy} 年失敗: {e3}")
    except Exception as e4:
        print(f"[ERROR] 備用方式失敗: {e4}")

def classify_type(market: str) -> str:
    m = market.strip()
    if "中央登錄公債" in m: return "bond"
    if "創新板" in m: return "tib"
    if any(k in m for k in ["初上市", "初上櫃", "第一上市", "第一上櫃"]): return "ipo"
    return "spo"

def safe_float(s) -> float:
    try: return float(str(s).replace(",", ""))
    except: return 0.0

def safe_int(s) -> int:
    try: return int(str(s).replace(",", ""))
    except: return 0

seen_bond_keys = set()
items = []
uid = 1

for cells in all_rows:
    # 判斷是 list（JSON格式欄位是純list）還是 HTML解析的 list
    if len(cells) < 17:
        continue

    # 欄位對應：
    # 0:序號 1:抽籤日 2:名稱 3:代號 4:市場 5:申購開始 6:申購結束
    # 7:承銷股數 8:實際承銷股數 9:承銷價 10:實際承銷價
    # 11:撥券/上市日 12:主辦券商 13:申購股數 14:總金額 15:合格件數 16:中籤率
    try:
        int(str(cells[0]).strip())
    except ValueError:
        continue

    market = str(cells[4]).strip()
    itype = classify_type(market)

    if itype == "bond":
        bkey = str(cells[2]) + str(cells[1])
        if bkey in seen_bond_keys:
            continue
        seen_bond_keys.add(bkey)

    item = {
        "id": uid,
        "name": str(cells[2]).strip(),
        "code": str(cells[3]).strip(),
        "type": itype,
        "market": market,
        "subStart": str(cells[5]).strip(),
        "subEnd": str(cells[6]).strip(),
        "drawDate": str(cells[1]).strip(),
        "listDate": str(cells[11]).strip(),
        "shares": safe_int(cells[8] if cells[8] else cells[7]),
        "price": str(cells[9]).strip(),
        "realPrice": str(cells[10]).strip(),
        "broker": str(cells[12]).strip(),
        "subShares": safe_int(cells[13]),
        "totalAmt": safe_int(cells[14]),
        "qualified": safe_int(cells[15]),
        "winRate": safe_float(cells[16]),
    }
    items.append(item)
    uid += 1

def roc_to_date(s: str):
    try:
        p = s.split("/")
        return date(int(p[0]) + 1911, int(p[1]), int(p[2]))
    except:
        return date(2000, 1, 1)

items.sort(key=lambda x: roc_to_date(x["drawDate"]), reverse=True)

output = {
    "updated": datetime.now(datetime.now().astimezone().tzinfo).strftime("%Y-%m-%dT%H:%M:%SZ"),
    "count": len(items),
    "data": items,
}

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"[OK] data.json 寫入完成，共 {len(items)} 筆")
if len(items) == 0:
    print("[ERROR] 資料為 0 筆！請檢查 TWSE URL 是否有效")
    exit(1)
