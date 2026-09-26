"""Alpha Lab のデータ更新（GitHub Actions で実行）

使い方: python alpha-lab/update.py <quotes|score|analyst|all> <出力フォルダ>
出力フォルダには前回のデータ（alpha-lab-data ブランチの中身）が入っている前提で、上書きして保存する。
外部ライブラリは使わない（標準ライブラリのみ）。
"""
import http.cookiejar, json, math, os, sys, time
import urllib.error, urllib.parse, urllib.request
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
NOW = datetime.now(timezone(timedelta(hours=9)))
AT = NOW.strftime("%Y-%m-%d %H:%M")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"
HD = {"User-Agent": UA, "Accept": "application/json, text/plain, */*", "Accept-Language": "en-US,en;q=0.9"}
NQ = {**HD, "Origin": "https://www.nasdaq.com", "Referer": "https://www.nasdaq.com/"}
Y = "https://query2.finance.yahoo.com"
CHUNK = 240000
OP = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
OP.addheaders = list(HD.items())
MACRO = ["^GSPC", "^IXIC", "^DJI", "^RUT", "^SOX", "^N225", "^VIX", "^VIX3M", "^MOVE",
         "^IRX", "2YY=F", "^FVX", "^TNX", "^TYX", "CL=F", "BZ=F", "NG=F", "GC=F", "SI=F", "HG=F",
         "JPY=X", "EURUSD=X", "DX-Y.NYB", "BTC-USD", "ETH-USD"]
FRED = {"DFF": 2, "DGS2": 2, "DGS10": 2, "T10Y2Y": 2, "BAMLH0A0HYM2": 2, "T10YIE": 2, "CPIAUCSL": 4, "CPILFESL": 4, "PCEPILFE": 4,
        "UNRATE": 4, "PAYEMS": 4, "ICSA": 2, "NFCI": 2, "MORTGAGE30US": 2, "UMCSENT": 4, "A191RL1Q225SBEA": 5,
        "RECPROUSM156N": 4, "WALCL": 3, "M2SL": 4}
SCOLS = ["roe", "opm", "pm", "gm", "rg", "eg", "pe", "fpe", "ps", "pb", "ev", "peg", "beta", "dy",
         "ma50", "ma200", "hi52", "lo52", "roa", "eps", "evr", "gp", "q"]
TYPES = ",".join([f"quarterly{k}" for k in ("TotalRevenue", "NetIncome", "DilutedEPS", "StockholdersEquity", "TotalAssets")]
                 + [f"annual{k}" for k in ("TotalRevenue", "NetIncome", "DilutedEPS")]
                 + [f"trailing{k}" for k in ("TotalRevenue", "NetIncome", "OperatingIncome", "GrossProfit", "DilutedEPS",
                                             "PeRatio", "ForwardPeRatio", "PegRatio", "PsRatio", "PbRatio",
                                             "EnterprisesValueEBITDARatio", "EnterprisesValueRevenueRatio")])


# ---------- 共通 ----------
def load(p, d):
    try:
        with open(p) as f:
            return json.load(f)
    except Exception:
        return d


def save(p, o):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w") as f:
        json.dump(o, f, ensure_ascii=False, separators=(",", ":"))


def http(url, headers=None, tries=3, raw=False):
    """GET して JSON（raw=True なら文字列）を返す。429 と一時的な失敗は待って再試行する"""
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=headers or HD)
            with OP.open(req, timeout=25) as r:
                b = r.read().decode("utf-8", "replace")
                return b if raw else json.loads(b)
        except urllib.error.HTTPError as e:
            if e.code in (401, 403, 404, 418) or i == tries - 1:
                raise
            time.sleep(4 * (i + 1) if e.code == 429 else 2 * (i + 1))
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(2 * (i + 1))


def why(e):
    return (f"http{e.code}" if isinstance(e, urllib.error.HTTPError)
            else (type(e).__name__ + " " + str(getattr(e, "reason", e))))[:60]


def attempt(f, *a):
    try:
        return f(*a)
    except Exception:
        return None


def num(x):
    if isinstance(x, dict):
        x = x.get("raw")
    if not isinstance(x, (int, float)) or isinstance(x, bool) or not math.isfinite(x):
        return None
    return int(round(x)) if abs(x) >= 1e5 else round(x, 4)


def r4(x):
    return f"{round(x, 4):.4f}".rstrip("0").rstrip(".")


def ysym(t):
    return t.replace(".", "-")


def pmap(f, items, n=6):
    with ThreadPoolExecutor(n) as ex:
        return list(ex.map(f, items))


def chunks(d, prefix):
    """{銘柄: 値} を 240KB ずつの塊に分ける → [(id, {銘柄: 値})]"""
    out, cur, size = [], {}, 0
    for t in sorted(d):
        s = len(json.dumps(d[t], ensure_ascii=False)) + len(t) + 6
        if cur and size + s > CHUNK:
            out.append(cur)
            cur, size = {}, 0
        cur[t] = d[t]
        size += s
    if cur:
        out.append(cur)
    return [(f"{prefix}{i:02d}", c) for i, c in enumerate(out)]


def crumb():
    """Yahoo Finance の認証用トークン（fc.yahoo.com でクッキーを受け取ってから取得）"""
    for pre in ("https://fc.yahoo.com/", "https://finance.yahoo.com/quote/AAPL/"):
        try:
            OP.open(urllib.request.Request(pre, headers=HD), timeout=15)
        except Exception:
            pass  # 404 でもクッキーは受け取れる
        for host in ("query2", "query1"):
            try:
                c = http(f"https://{host}.finance.yahoo.com/v1/test/getcrumb", tries=1, raw=True).strip()
                if c and "<" not in c and len(c) < 40:
                    return c
            except Exception:
                pass
    return None


# ---------- 株価・チャート ----------
def chart(t):
    j = http(f"{Y}/v8/finance/chart/{ysym(t)}?range=1y&interval=1d")
    r = ((j.get("chart") or {}).get("result") or [None])[0]
    if not r:
        raise ValueError("no_data")
    m, q = r["meta"], r["indicators"]["quote"][0]
    off = m.get("gmtoffset", -14400)
    bars = {}
    for i, ts in enumerate(r.get("timestamp") or []):
        o, h, l, c, v = (q.get(k, [None])[i] for k in ("open", "high", "low", "close", "volume"))
        if None in (o, h, l, c) or c <= 0:
            continue
        bars[datetime.fromtimestamp(ts + off, timezone.utc).strftime("%Y-%m-%d")] = (o, max(h, o, c), min(l, o, c), c, v or 0)
    if not bars:
        raise ValueError("no_bars")
    today = datetime.fromtimestamp(time.time() + off, timezone.utc).strftime("%Y-%m-%d")
    end = ((m.get("currentTradingPeriod") or {}).get("regular") or {}).get("end", 0)
    live = max(bars) == today and time.time() < end  # 取引中の当日足は未確定なのでチャートに入れない
    px = m.get("regularMarketPrice") or bars[max(bars)][3]
    pt = datetime.fromtimestamp((m.get("regularMarketTime") or time.time()) + off, timezone.utc).strftime("%Y-%m-%d %H:%M")
    if live:
        bars.pop(today)
    return bars, px, pt, live


def macro(sym):
    j = http(f"{Y}/v8/finance/chart/{urllib.parse.quote(sym, safe='')}?range=1y&interval=1d")
    r = ((j.get("chart") or {}).get("result") or [None])[0]
    if not r:
        raise ValueError("no_data")
    m, q = r["meta"], (r["indicators"]["quote"] or [{}])[0]
    off, t, c = m.get("gmtoffset", 0), [], []
    for i, ts in enumerate(r.get("timestamp") or []):
        v = (q.get("close") or [None] * (i + 1))[i]
        if v is None or v <= 0:
            continue
        d = (ts + off) // 86400
        if t and t[-1] == d:
            c[-1] = v
        else:
            t.append(d)
            c.append(v)
    if len(c) < 5:
        raise ValueError("no_bars")
    sig = lambda x: float(f"{x:.5g}")
    return {"t": t, "c": [sig(x) for x in c], "p": sig(m.get("regularMarketPrice") or c[-1]), "tm": m.get("regularMarketTime")}


def news(t):
    s = ysym(t)
    j = http(f"{Y}/v1/finance/search?q={urllib.parse.quote(s)}&quotesCount=0&newsCount=8&enableFuzzyQuery=false")
    out = []
    for n in j.get("news") or []:
        rel = n.get("relatedTickers") or []
        if n.get("title") and (not rel or s in rel or t in rel):
            out.append([n["title"][:170], (n.get("publisher") or "")[:40], n.get("providerPublishTime") or 0, (n.get("link") or "")[:300]])
        if len(out) == 5:
            break
    return out


def fred(sid, years):
    start = (NOW - timedelta(days=365 * years)).strftime("%Y-%m-%d")
    txt = http(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}&cosd={start}", raw=True)
    t, c = [], []
    for line in txt.strip().splitlines()[1:]:
        d, _, v = line.partition(",")
        try:  # 欠損（空欄や "."）の日は飛ばす。日付と値は両方そろったときだけ入れる
            day = int(datetime.strptime(d.strip(), "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp() // 86400)
            val = float(f"{float(v):.5g}")
        except ValueError:
            continue
        t.append(day)
        c.append(val)
    if not c:
        raise ValueError("no_data")
    return {"t": t, "c": c}


def fng():
    j = http(f"https://production.dataviz.cnn.io/index/fearandgreed/graphdata/{(NOW - timedelta(days=370)).strftime('%Y-%m-%d')}",
             headers={**HD, "Origin": "https://edition.cnn.com", "Referer": "https://edition.cnn.com/"})
    f = j["fear_and_greed"]
    r1 = lambda v: round(v, 1) if isinstance(v, (int, float)) else None
    keys = ("market_momentum_sp500", "stock_price_strength", "stock_price_breadth", "put_call_options",
            "market_volatility_vix", "junk_bond_demand", "safe_haven_demand")
    return {"s": r1(f.get("score")), "r": f.get("rating", ""), "pc": r1(f.get("previous_close")), "w1": r1(f.get("previous_1_week")),
            "m1": r1(f.get("previous_1_month")), "y1": r1(f.get("previous_1_year")), "ts": f.get("timestamp", ""),
            "h": [[int(p["x"] // 86400000), r1(p["y"])] for p in (j.get("fear_and_greed_historical") or {}).get("data") or []],
            "c": {k: [r1((j.get(k) or {}).get("score")), (j.get(k) or {}).get("rating", "")] for k in keys if j.get(k)}}


def calendar(universe):
    days = [d for d in (NOW.date() + timedelta(days=i) for i in range(-4, 15)) if d.weekday() < 5]
    earn, econ = [], []
    for d in days:
        ds = d.strftime("%Y-%m-%d")
        if d >= NOW.date():
            rows = ((http(f"https://api.nasdaq.com/api/calendar/earnings?date={ds}", NQ).get("data") or {}).get("rows") or [])
            earn += [[ds, r["symbol"].replace("/", "."), r.get("time", ""), r.get("epsForecast", ""), r.get("noOfEsts", ""),
                      r.get("lastYearEPS", ""), r.get("fiscalQuarterEnding", "")]
                     for r in rows if (r.get("symbol") or "").replace("/", ".") in universe]
        if d <= NOW.date() + timedelta(days=7):
            rows = ((http(f"https://api.nasdaq.com/api/calendar/economicevents?date={ds}", NQ).get("data") or {}).get("rows") or [])
            econ += [[ds, r.get("gmt", ""), r.get("eventName", ""), r.get("actual", ""), r.get("consensus", ""), r.get("previous", "")]
                     for r in rows if r.get("country") == "United States" and r.get("eventName")]
    return earn, econ


def run_quotes(out, tickers, log):
    t0 = time.time()
    res, err = {}, {}
    for t, r in zip(tickers, pmap(lambda t: attempt(chart, t), tickers)):
        if r:
            res[t] = r
        else:
            err[t] = "chart_missing"
    if len(res) < len(tickers) * 0.8:
        try:
            chart("AAPL")
            why_s = "?"
        except Exception as e:
            why_s = why(e)
        log.append(["bulk", "error", f"株価・チャートを更新できず: Yahoo Finance から取得できたのが {len(res)}/{len(tickers)}社のみ（{why_s}）"])
        return
    # チャート（1年分）
    rows = {t: ";".join(d + "," + ",".join([*(r4(x) for x in b[:4]), str(int(b[4]))]) for d, b in sorted(bars.items()))
            for t, (bars, p, pt, live) in res.items()}
    cmap = {}
    for cid, ch in chunks(rows, "c"):
        save(f"{out}/bars/{cid}.json", {"at": AT, "d": ch})
        cmap.update({t: cid for t in ch})
    # 株価: 時価総額は株数据え置きで株価に比例、PER は利益据え置きで株価に比例
    q = load(f"{out}/quotes.json", None) or load(f"{HERE}/seed/quotes.json", {})
    old_rows, old_pe = q.get("rows", {}), q.get("pe", {})
    new_rows, new_pe, live_n, asof = {}, {}, 0, ""
    for t, (bars, p, pt, live) in res.items():
        old = old_rows.get(t) or [None, None, None, None]
        y0 = [bars[d][3] for d in sorted(bars) if d < pt[:4] + "-01-01"]
        k = p / old[1] if old[1] else None
        mc = old[0] * k if k and old[0] else old[0]
        new_rows[t] = [round(mc, 2 if mc < 100 else 1) if mc else mc, round(p, 2), old[2], round(p / y0[-1] - 1, 4) if y0 else old[3]]
        if k and old_pe.get(t):
            new_pe[t] = round(old_pe[t] * k, 1)
        live_n += 1 if live else 0
        asof = max(asof, pt[:10])
    for t in old_rows:  # 取れなかった銘柄は前回の値を残す
        new_rows.setdefault(t, old_rows[t])
    save(f"{out}/quotes.json", {"rows": new_rows, "pe": {**old_pe, **new_pe}, "asOf": asof, "updatedAt": AT})
    # マーケット指標
    mac = dict(zip(MACRO, pmap(lambda s: attempt(macro, s), MACRO)))
    merr = [s for s, v in mac.items() if not v]
    mac = {s: v for s, v in mac.items() if v}
    if mac:
        save(f"{out}/macro.json", {"at": AT, "items": mac, "err": merr})
    # ニュース
    nw = dict(zip(tickers, pmap(lambda t: attempt(news, t), tickers, 8)))
    nw = {t: v for t, v in nw.items() if v}
    nmap = {}
    for nid, ch in chunks(nw, "n"):
        save(f"{out}/news/{nid}.json", {"at": AT, "d": ch})
        nmap.update({t: nid for t in ch})
    # 経済データ・Fear & Greed・話題の銘柄
    diag, old_econ = {}, load(f"{out}/econ.json", {})
    econ = {"at": AT, "fred": {}, "fng": None, "trend": [], "diag": diag}
    try:
        q = (http(f"{Y}/v1/finance/trending/US?count=25").get("finance") or {}).get("result") or [{}]
        econ["trend"] = [x["symbol"] for x in (q[0].get("quotes") or []) if x.get("symbol")]
    except Exception as e:
        diag["trend"] = why(e)
    try:
        fred("DFF", 1)
        econ["fred"] = {s: v for s, v in zip(FRED, pmap(lambda s: attempt(fred, s, FRED[s]), list(FRED))) if v}
    except Exception as e:
        diag["fred"] = why(e)
    try:
        econ["fng"] = fng()
    except Exception as e:
        diag["fng"] = why(e)
    if econ["fred"] or econ["fng"]:
        econ["fat"] = AT
    econ["fred"] = econ["fred"] or old_econ.get("fred") or {}
    econ["fng"] = econ["fng"] or old_econ.get("fng")
    econ.setdefault("fat", old_econ.get("fat"))
    save(f"{out}/econ.json", econ)
    # 決算と経済指標の予定
    try:
        earn, ev = calendar(set(tickers))
        save(f"{out}/cal.json", {"at": AT, "earn": earn, "econ": ev})
    except Exception as e:
        diag["cal"] = why(e)
        save(f"{out}/econ.json", econ)
    sec = round(time.time() - t0)
    note = (f"一斉更新: 株価{len(res)}社（{asof}{'・取引中の値を含む' if live_n else ''}）、チャート日足1年分{len(rows)}社、"
            f"マーケット指標{len(mac)}種、ニュース{len(nw)}社、経済データ{len(econ['fred'])}系列"
            f"（GitHub Actions、{sec}秒）" + (f"、取得できず {len(err)}社" if err else "") + (f"、接続できず {', '.join(diag)}" if diag else ""))
    save(f"{out}/bulk.json", {"at": AT, "day": NOW.strftime("%Y-%m-%d"), "asof": asof,
                              "last": max((max(b) for b, *_ in res.values() if b), default=""), "src": "Yahoo Finance",
                              "n": len(res), "live": live_n, "err": err, "note": note, "map": cmap, "chunks": len(set(cmap.values())),
                              "nmap": nmap, "via": "github"})
    log.append(["bulk", "partial" if err or diag else "ok", note])


# ---------- スコア用の財務指標 ----------
def days_between(a, b):
    return (date.fromisoformat(b) - date.fromisoformat(a)).days


def timeseries(t, cr):
    s = ysym(t)
    j = http(f"{Y}/ws/fundamentals-timeseries/v1/finance/timeseries/{s}?symbol={s}&type={TYPES}"
             f"&period1={int(time.time()) - 3 * 366 * 86400}&period2={int(time.time()) + 86400}"
             + (f"&crumb={urllib.parse.quote(cr)}" if cr else ""))
    out = {}
    for r in (j.get("timeseries") or {}).get("result") or []:
        k = ((r.get("meta") or {}).get("type") or [None])[0]
        pts = sorted((x["asOfDate"], num(x.get("reportedValue"))) for x in (r.get(k) or [])
                     if isinstance(x, dict) and x.get("asOfDate") and num(x.get("reportedValue")) is not None)
        if k and pts:
            out[k] = pts
    if not any(k.startswith(("quarterly", "annual")) for k in out):
        raise LookupError("no_financials")
    return out


def price_hist(t):
    j = http(f"{Y}/v8/finance/chart/{ysym(t)}?range=1y&interval=1d&events=div")
    r = ((j.get("chart") or {}).get("result") or [None])[0]
    if not r:
        raise LookupError("no_chart")
    q = (r.get("indicators") or {}).get("quote", [{}])[0]
    cl, hi, lo = {}, [], []
    for i, ts in enumerate(r.get("timestamp") or []):
        c, h, l = (q.get(k, [None] * (i + 1))[i] for k in ("close", "high", "low"))
        if c and c > 0:
            cl[ts // 86400] = c
            hi.append(h or c)
            lo.append(l or c)
    dv = sum(d.get("amount") or 0 for d in ((r.get("events") or {}).get("dividends") or {}).values()
             if (d.get("date") or 0) > time.time() - 365 * 86400)
    return {"cl": cl, "hi": max(hi) if hi else None, "lo": min(lo) if lo else None, "dv": dv}


def rets(cl):
    ks = sorted(cl)
    return {b: cl[b] / cl[a] - 1 for a, b in zip(ks, ks[1:])}


def beta(cl, mk):
    r = rets(cl)
    ds = [d for d in r if d in mk]
    if len(ds) < 60:
        return None
    x, y = [mk[d] for d in ds], [r[d] for d in ds]
    mx, my = sum(x) / len(x), sum(y) / len(y)
    vx = sum((a - mx) ** 2 for a in x)
    return sum((a - mx) * (b - my) for a, b in zip(x, y)) / vx if vx else None


def score_row(ts, ch, mk):
    last = lambda k: ts[k][-1][1] if ts.get(k) else None

    def ttm(k):
        v = last("trailing" + k)
        q = ts.get("quarterly" + k) or []
        if v is None and len(q) >= 4 and days_between(q[-4][0], q[-1][0]) < 300:
            v = sum(x[1] for x in q[-4:])
        return v

    def yoy(k):
        q = ts.get("quarterly" + k) or []
        for d0, v0 in q[:-1]:
            if 330 <= days_between(d0, q[-1][0]) <= 400:
                return q[-1][1] / v0 - 1 if v0 > 0 else None
        a = ts.get("annual" + k) or []
        return a[-1][1] / a[-2][1] - 1 if len(a) >= 2 and a[-2][1] > 0 else None

    div = lambda a, b: a / b if a is not None and b else None
    rev, ni, eq, ta = ttm("TotalRevenue"), ttm("NetIncome"), last("quarterlyStockholdersEquity"), last("quarterlyTotalAssets")
    rev = rev if rev and rev > 0 else None
    eg = yoy("DilutedEPS")
    row = {"roe": div(ni, eq) if eq and eq > 0 else None, "opm": div(ttm("OperatingIncome"), rev), "pm": div(ni, rev),
           "gm": div(ttm("GrossProfit"), rev), "rg": yoy("TotalRevenue"), "eg": eg if eg is not None else yoy("NetIncome"),
           "pe": last("trailingPeRatio"), "fpe": last("trailingForwardPeRatio"), "ps": last("trailingPsRatio"),
           "pb": last("trailingPbRatio"), "ev": last("trailingEnterprisesValueEBITDARatio"), "peg": last("trailingPegRatio"),
           "roa": div(ni, ta), "eps": ttm("DilutedEPS"), "evr": last("trailingEnterprisesValueRevenueRatio"),
           "gp": ttm("GrossProfit"), "q": max((v[-1][0] for k, v in ts.items() if k.startswith("quarterly")), default=None)}
    if ch:
        c = [ch["cl"][d] for d in sorted(ch["cl"])]
        row.update({"beta": beta(ch["cl"], mk) if mk else None, "dy": ch["dv"] / c[-1] if c else None,
                    "ma50": sum(c[-50:]) / 50 if len(c) >= 50 else None,
                    "ma200": sum(c[-200:]) / 200 if len(c) >= 200 else None, "hi52": ch["hi"], "lo52": ch["lo"]})
    return [row.get(k) if k == "q" else num(row.get(k)) for k in SCOLS]


def run_score(out, tickers, cr, log):
    t0 = time.time()
    mk = attempt(lambda: rets(price_hist("SPY")["cl"]))

    def one(t):
        try:
            return score_row(timeseries(t, cr), attempt(price_hist, t), mk), None
        except Exception as e:
            return None, str(e)[:40]

    rows, err = {}, {}
    for t, (row, e) in zip(tickers, pmap(one, tickers)):
        if row:
            rows[t] = row
        else:
            err[t] = e
    if len(rows) < len(tickers) * 0.6:
        log.append(["score", "error", f"スコア用の財務指標を更新できず: 取得できたのが {len(rows)}/{len(tickers)}社のみ"
                    f"（主な原因 {Counter(err.values()).most_common(1)[0][0] if err else '?'}）"])
        return
    prev = load(f"{out}/fund.json", {}).get("rows", {})
    save(f"{out}/fund.json", {"at": AT, "src": "Yahoo Finance", "n": len(rows), "err": err, "cols": SCOLS,
                              "rows": {**{t: v for t, v in prev.items() if t in err}, **rows}})
    log.append(["score", "partial" if err else "ok",
                f"スコア用の財務指標 {len(rows)}社（GitHub Actions、{round(time.time() - t0)}秒）" + (f"、取得できず {len(err)}社" if err else "")])


# ---------- アナリスト評価 ----------
def analyst_row(t, cr):
    j = http(f"{Y}/v10/finance/quoteSummary/{ysym(t)}?modules=financialData,recommendationTrend,upgradeDowngradeHistory"
             f"&crumb={urllib.parse.quote(cr)}")
    r = ((j.get("quoteSummary") or {}).get("result") or [None])[0]
    if not r:
        raise LookupError("no_result")
    f = r.get("financialData") or {}
    tr = {x.get("period"): [num(x.get(k)) or 0 for k in ("strongBuy", "buy", "hold", "sell", "strongSell")]
          for x in (r.get("recommendationTrend") or {}).get("trend") or []}
    ud = []
    for h in sorted((r.get("upgradeDowngradeHistory") or {}).get("history") or [],
                    key=lambda h: h.get("epochGradeDate") or 0, reverse=True)[:3]:
        d = datetime.fromtimestamp(h.get("epochGradeDate") or 0, timezone.utc).strftime("%Y-%m-%d")
        ud.append([d, h.get("firm") or "", h.get("action") or "", h.get("fromGrade") or "", h.get("toGrade") or "",
                   num(h.get("currentPriceTarget"))])
    row = {"m": num(f.get("recommendationMean")), "k": f.get("recommendationKey") or "",
           "n": num(f.get("numberOfAnalystOpinions")), "p": num(f.get("currentPrice")),
           "t": [num(f.get(k)) for k in ("targetMeanPrice", "targetHighPrice", "targetLowPrice")],
           "tr": tr.get("0m"), "tr1": tr.get("-1m"), "ud": ud}
    if not (row["m"] or any(row["tr"] or []) or ud):
        raise LookupError("no_coverage")
    return row


def insights_row(t):
    j = http(f"{Y}/ws/insights/v2/finance/insights?symbol={ysym(t)}&lang=en-US&region=US")
    r = (j.get("finance") or {}).get("result") or {}
    rec = r.get("recommendation") or {}
    rep = sorted((x for x in r.get("reports") or [] if isinstance(x, dict) and x.get("investmentRating")),
                 key=lambda x: x.get("reportDate") or "", reverse=True)
    val = (r.get("instrumentInfo") or {}).get("valuation") or {}
    row = {"pr": (rec.get("rating") or "").upper(), "pv": rec.get("provider") or "", "tp": num(rec.get("targetPrice")),
           "rr": [(rep[0].get("reportDate") or "")[:10], rep[0].get("provider") or "", rep[0].get("investmentRating") or "",
                  num(rep[0].get("targetPrice")), rep[0].get("targetPriceStatus") or ""] if rep else None,
           "va": val.get("description") or ""}
    if not (row["pr"] or row["rr"]):
        raise LookupError("no_coverage")
    return row


def run_analyst(out, tickers, cr, log):
    """コンセンサス（要トークン）に、調査会社（Argus など）の評価を重ねる"""
    t0 = time.time()

    def one(t):
        row, errs = {}, []
        if cr:
            try:
                row = analyst_row(t, cr)
            except Exception as e:
                errs.append(str(e)[:30])
        try:
            ins = insights_row(t)
            if not row:
                row = {"m": None, "k": "", "n": None, "p": None, "t": [ins["tp"], None, None], "tr": None, "tr1": None, "ud": []}
            row.update({k: v for k, v in ins.items() if k != "tp"})
        except Exception as e:
            errs.append(str(e)[:30])
        return (row or None), (errs[0] if errs and not row else None)

    rows, err = {}, {}
    for t, (row, e) in zip(tickers, pmap(one, tickers, 4)):
        if row:
            rows[t] = row
        else:
            err[t] = e or "no_coverage"
    if len(rows) < len(tickers) * 0.3:
        log.append(["analyst", "error", f"アナリスト評価を更新できず: 取得できたのが {len(rows)}/{len(tickers)}社のみ"
                    f"（トークン{'あり' if cr else 'なし'}、主な原因 {Counter(err.values()).most_common(1)[0][0] if err else '?'}）"])
        return
    prev = load(f"{out}/analyst.json", {}).get("rows", {})
    src = "Yahoo Finance" if cr else "Yahoo Finance（調査会社の評価）"
    save(f"{out}/analyst.json", {"at": AT, "src": src, "n": len(rows), "err": err,
                                 "rows": {**{t: v for t, v in prev.items() if t in err}, **rows}})
    log.append(["analyst", "partial" if err else "ok", f"アナリスト評価 {len(rows)}社（{src}、GitHub Actions、{round(time.time() - t0)}秒）"
                + (f"、取得できず {len(err)}社" if err else "") + ("" if cr else "、コンセンサスは取得できず")])


# ---------- 調査用（応答の形を確かめる） ----------
def run_probe(out, log):
    cr = crumb()
    res = {"crumb": bool(cr)}
    q = urllib.parse.quote(cr or "")
    for t in ("AAPL", "JPM", "NVDA"):
        for mods in ("calendarEvents,earningsTrend,earningsHistory", "defaultKeyStatistics"):
            try:
                res[f"{t}:{mods}"] = http(f"{Y}/v10/finance/quoteSummary/{t}?modules={mods}&crumb={q}")
            except Exception as e:
                res[f"{t}:{mods}"] = why(e)
        try:
            res[f"{t}:chart1d"] = http(f"{Y}/v8/finance/chart/{t}?range=1d&interval=1d")["chart"]["result"][0]["meta"]
        except Exception as e:
            res[f"{t}:chart1d"] = why(e)
    start, end = NOW.strftime("%Y-%m-%d"), (NOW + timedelta(days=10)).strftime("%Y-%m-%d")
    bodies = {
        "econ": {"sortType": "ASC", "entityIdType": "economic_event", "sortField": "startdatetime",
                 "includeFields": ["econ_release", "country_code", "startdatetime", "period", "after_release_actual",
                                   "consensus_estimate", "prior_release_actual", "originally_reported_actual"],
                 "query": {"operator": "and", "operands": [{"operator": "gte", "operands": ["startdatetime", start]},
                                                           {"operator": "lte", "operands": ["startdatetime", end]}]},
                 "offset": 0, "size": 100},
        "earn": {"sortType": "DESC", "entityIdType": "sp_earnings", "sortField": "intradaymarketcap",
                 "includeFields": ["ticker", "companyshortname", "intradaymarketcap", "eventname", "startdatetime",
                                   "startdatetimetype", "epsestimate", "epsactual", "epssurprisepct"],
                 "query": {"operator": "and", "operands": [{"operator": "gte", "operands": ["startdatetime", start]},
                                                           {"operator": "lt", "operands": ["startdatetime", end]},
                                                           {"operator": "eq", "operands": ["region", "us"]}]},
                 "offset": 0, "size": 100},
    }
    for k, b in bodies.items():
        for host in ("query1", "query2"):
            try:
                req = urllib.request.Request(f"https://{host}.finance.yahoo.com/v1/finance/visualization?lang=en-US&region=US&crumb={q}",
                                             data=json.dumps(b).encode(), headers={**HD, "Content-Type": "application/json"})
                with OP.open(req, timeout=25) as r:
                    res[f"viz:{k}:{host}"] = json.loads(r.read().decode("utf-8", "replace"))
            except Exception as e:
                res[f"viz:{k}:{host}"] = why(e)
    save(f"{out}/probe/probe.json", res)
    log.append(["probe", "ok", f"調査 {len(res)}件（トークン{'あり' if cr else 'なし'}）"])


# ---------- 入口 ----------
def main():
    mode = (sys.argv[1] if len(sys.argv) > 1 else "all").strip() or "all"
    out = sys.argv[2] if len(sys.argv) > 2 else "out"
    os.makedirs(out, exist_ok=True)
    q = load(f"{out}/quotes.json", None) or load(f"{HERE}/seed/quotes.json", {})
    tickers = sorted(q.get("rows", {}))
    log = []
    if mode == "probe":
        run_probe(out, log)
        print(log)
        return
    if mode in ("quotes", "all"):
        run_quotes(out, tickers, log)
    cr = crumb() if mode in ("score", "analyst", "all") else None
    if mode in ("score", "all"):
        run_score(out, tickers, cr, log)
    if mode in ("analyst", "all"):
        run_analyst(out, tickers, cr, log)
    runs = load(f"{out}/runs.json", {}).get("items", [])
    runs = sorted(runs + [{"at": AT, "job": j, "status": s, "note": n} for j, s, n in log], key=lambda x: str(x.get("at", "")))[-80:]
    save(f"{out}/runs.json", {"items": runs})
    save(f"{out}/status.json", {"at": AT, "mode": mode, "log": log})
    for j, s, n in log:
        print(f"[{j}] {s}: {n}")


if __name__ == "__main__":
    main()
