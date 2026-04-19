#!/usr/bin/env python3
"""
每天自動從 TWSE 抓公開申購資料，輸出 data.json 供網頁使用。
"""
import json, re, urllib.request
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
    try:
        raw = fetch_url(f"https://www.twse.com.tw/announcement/publicForm?response=json&yy={yy}")
        data = json.loads(raw)
        rows = data.get("data", [])
        print(f"[INFO] {yy} 年 JSON：{len(rows)} 筆")
        all_rows.extend(rows)
        continue
    except Exception as e:
        print(f"[WARN] {yy} 年 JSON 失敗: {e}，改 HTML")
    try:
        html = fetch_url(f"https://www.twse.com.tw/announcement/publicForm?response=html&yy={yy}")
        count = 0
        for row in re.findall(r"<tr>(.*?)</tr>", html, re.DOTALL):
            cells = [re.sub(r"<[^>]+>", "", c).strip()
                     for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)]
            if len(cells) >= 17:
                try:
                    int(cells[0])
                    all_rows.append(cells)
                    count += 1
                except ValueError:
                    pass
        print(f"[INFO] {yy} 年 HTML：{count} 筆")
    except Exception as e2:
        print(f"[ERROR] {yy} 年失敗: {e2}")

print(f"[INFO] 共抓到 {len(all_rows)} 筆原始資料")

def classify_type(market):
    m = market.strip()
    if "中央登錄公債" in m: return "bond"
    if "創新板" in m: return "tib"
    if any(k in m for k in ["初上市", "初上櫃", "第一上市", "第一上櫃"]): return "ipo"
    return "spo"

def safe_float(s):
    try: return float(str(s).replace(",", ""))
    except: return 0.0

def safe_int(s):
    try: return int(str(s).replace(",", ""))
    except: return 0

seen_keys = set()
items = []
uid = 1

for cells in all_rows:
    if len(cells) < 17:
        continue
    try:
        int(str(cells[0]).strip())
    except ValueError:
        continue

    market = str(cells[4]).strip()
    itype  = classify_type(market)
    name   = str(cells[2]).strip()
    code   = str(cells[3]).strip()
    draw   = str(cells[1]).strip()

    dedup_key = (name + "|" + draw) if itype == "bond" else (code + "|" + draw)
    if dedup_key in seen_keys:
        continue
    seen_keys.add(dedup_key)

    items.append({
        "id": uid,
        "name": name,
        "code": code,
        "type": itype,
        "market": market,
        "subStart":  str(cells[5]).strip(),
        "subEnd":    str(cells[6]).strip(),
        "drawDate":  draw,
        "listDate":  str(cells[11]).strip(),
        "shares":    safe_int(cells[8] or cells[7]),
        "price":     str(cells[9]).strip(),
        "realPrice": str(cells[10]).strip(),
        "broker":    str(cells[12]).strip(),
        "subShares": safe_int(cells[13]),
        "totalAmt":  safe_int(cells[14]),
        "qualified": safe_int(cells[15]),
        "winRate":   safe_float(cells[16]),
    })
    uid += 1

def roc_to_date(s):
    try:
        p = s.split("/")
        return date(int(p[0]) + 1911, int(p[1]), int(p[2]))
    except:
        return date(2000, 1, 1)

items.sort(key=lambda x: roc_to_date(x["drawDate"]), reverse=True)

output = {
    "updated": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
    "count": len(items),
    "data": items,
}

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"[OK] data.json 寫入完成，共 {len(items)} 筆")
if len(items) == 0:
    print("[ERROR] 資料 0 筆！")
    exit(1)
