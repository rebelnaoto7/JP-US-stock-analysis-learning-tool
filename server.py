#!/usr/bin/env python3
"""
プロ投資家の視点 日米株アナライザー — ローカルサーバー

ブラウザのアプリから Yahoo Finance などへ直接アクセスすることはできない（CORS制限）ため、
このサーバーが代わりに検索・データ取得を行います。自分のPC（127.0.0.1）の中だけで動きます。

  起動：同じフォルダの「アプリを起動.bat」をダブルクリック（または python server.py）
  終了：この黒い画面を閉じる

  /api/search?q=   銘柄検索（JPX「東証上場銘柄一覧」＋ Yahoo Finance の検索）
  /api/stock?symbol=7203.T  Yahoo Finance（yfinance）から財務データを取得
  /api/market      日本（TOPIX連動ETF 1306）・米国（S&P 500）の1年騰落率と相場の方向

東証上場銘柄一覧：このフォルダに jp_stocks.json があればそれを使い、無ければ
このPCの %LOCALAPPDATA%\\ProInvestorAnalyzer に取得・保存します（共有フォルダには作りません）。

オプション：
  python server.py --update-list   東証上場銘柄一覧（JPX）を最新にする
  python server.py --port 8723 --no-browser
"""
import argparse
import datetime as dt
import io
import json
import math
import os
import pathlib
import re
import sys
import threading
import time
import unicodedata
import urllib.parse
import urllib.request
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

HERE = pathlib.Path(__file__).resolve().parent
APP_HTML = "プロ投資家の視点_日米株アナライザー.html"
# 東証上場銘柄一覧：アプリのフォルダに jp_stocks.json があればそれを使う。
# 無ければ、このPCの中（%LOCALAPPDATA%\ProInvestorAnalyzer）に取得・保存する（共有フォルダには作らない）
APP_LIST_FILE = HERE / "jp_stocks.json"
_LOCAL_BASE = os.environ.get("LOCALAPPDATA")
CACHE_LIST_FILE = (pathlib.Path(_LOCAL_BASE) / "ProInvestorAnalyzer" if _LOCAL_BASE else pathlib.Path.home() / ".pro_investor_analyzer") / "jp_stocks.json"
JPX_PAGE = "https://www.jpx.co.jp/markets/statistics-equities/misc/01.html"
JPX_FALLBACK = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx"
DEFAULT_PORT = 8723
VERSION = "1.0"

def _ensure_libs():
    """yfinance・pandas が無ければ、確認のうえ pip でインストールする。"""
    try:
        import pandas  # noqa: F401
        import yfinance  # noqa: F401
        import openpyxl  # noqa: F401  (JPXの銘柄一覧（xlsx）を読むのに必要)
        return
    except ImportError:
        pass
    print("必要なライブラリ（yfinance・pandas・openpyxl）が見つかりません。")
    try:
        ans = input("今すぐインストールしますか？（インターネット接続が必要です） [Y/n]：").strip().lower()
    except EOFError:
        ans = "n"
    if ans in ("", "y", "yes"):
        import importlib
        import site
        import subprocess
        rc = subprocess.call([sys.executable, "-m", "pip", "install", "yfinance", "pandas", "openpyxl"])
        try:
            site.addsitedir(site.getusersitepackages())
        except Exception:
            pass
        importlib.invalidate_caches()
        if rc == 0:
            print("インストールが完了しました。アプリを起動します。")
            return
    print("次のコマンドでインストールしてから、もう一度起動してください：")
    print("  python -m pip install yfinance pandas openpyxl")
    try:
        input("Enterキーで終了します…")
    except EOFError:
        pass
    sys.exit(1)


_ensure_libs()
try:
    import pandas as pd
    import yfinance as yf
except ImportError:
    print("ライブラリを読み込めませんでした。この画面を閉じて、もう一度「アプリを起動.bat」を開いてください。")
    try:
        input("Enterキーで終了します…")
    except EOFError:
        pass
    sys.exit(1)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) or math.isinf(f) else f


def mil(v):
    return None if v is None else round(v / 1e6, 3)


def r2(v, d=4):
    return None if v is None else round(v, d)


def norm(s):
    """全角/半角・大文字/小文字・ひらがな/カタカナの違いを吸収して比較用にする"""
    s = unicodedata.normalize("NFKC", str(s or "")).lower()
    s = "".join(chr(ord(c) + 0x60) if "ぁ" <= c <= "ゖ" else c for c in s)
    return re.sub(r"[\s・･\-‐－—\.,、。()（）&＆'’]", "", s)


def series(df, *names):
    """財務諸表の行を [(日付, 値), ...]（新しい順）で返す"""
    if df is None or getattr(df, "empty", True):
        return []
    for n in names:
        if n in df.index:
            return [(pd.Timestamp(c).tz_localize(None) if pd.Timestamp(c).tzinfo else pd.Timestamp(c), num(v)) for c, v in zip(df.columns, df.loc[n].values)]
    return []


def valid(pairs):
    return [(d, v) for d, v in pairs if v is not None]


def first(pairs, i=0):
    vs = valid(pairs)
    return vs[i][1] if len(vs) > i else None


class TTLCache:
    def __init__(self, ttl):
        self.ttl, self.d, self.lock = ttl, {}, threading.Lock()

    def get(self, k):
        with self.lock:
            v = self.d.get(k)
            if v and time.time() - v[0] < self.ttl:
                return v[1]
            return None

    def set(self, k, val):
        with self.lock:
            self.d[k] = (time.time(), val)


STOCK_CACHE, SEARCH_CACHE, MARKET_CACHE = TTLCache(15 * 60), TTLCache(10 * 60), TTLCache(60 * 60)
FETCH_SEM = threading.Semaphore(2)


# ---------------------------------------------------------------------------
# 東証上場銘柄一覧（JPX）
# ---------------------------------------------------------------------------
JP_LIST = {"date": "", "items": []}
JP_BY_CODE = {}


def jp_list_file():
    return APP_LIST_FILE if APP_LIST_FILE.exists() else CACHE_LIST_FILE


def download_jp_list():
    req = urllib.request.Request(JPX_PAGE, headers={"User-Agent": "Mozilla/5.0"})
    url = JPX_FALLBACK
    try:
        html = urllib.request.urlopen(req, timeout=30).read().decode("utf-8", "replace")
        m = re.search(r'href="([^"]+data_j\.xlsx?)"', html)
        if m:
            url = urllib.parse.urljoin(JPX_PAGE, m.group(1))
    except Exception:
        pass
    data = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=60).read()
    df = pd.read_excel(io.BytesIO(data), dtype=str)
    cols = list(df.columns)
    date = str(df.iloc[0, 0]) if len(df) else ""
    items = []
    for _, row in df.iterrows():
        seg = str(row[cols[3]])
        if "株式" not in seg:
            continue
        code = str(row[cols[1]]).strip()
        size = str(row[cols[8]]).strip() if len(cols) > 8 else ""
        items.append([code, str(row[cols[2]]).strip(), seg, str(row[cols[5]]).strip(), size if size not in ("-", "nan") else ""])
    out = {"date": date, "source": url, "items": items}
    dest = jp_list_file()
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    except OSError as e:
        print(f"  銘柄一覧を保存できませんでした（次回の起動でもう一度取得します）：{e}")
    return out


# TOPIXの規模区分（1=Core30, 2=Large70, 4=Mid400, 6=Small1, 7=Small2）→ 検索結果の並び順に使う
SIZE_RANK = {"1": 0, "2": 1, "4": 2, "6": 3, "7": 4}


def load_jp_list(force=False):
    global JP_LIST, JP_BY_CODE
    data = None
    src = jp_list_file()
    if src.exists() and not force:
        try:
            data = json.loads(src.read_text(encoding="utf-8"))
        except Exception:
            data = None
    if data is None:
        try:
            print("東証上場銘柄一覧（JPX）をダウンロードしています…")
            data = download_jp_list()
        except Exception as e:  # noqa: BLE001
            print(f"  銘柄一覧を取得できませんでした（日本語名での検索が使えません）：{e}")
            data = {"date": "", "items": []}
    for it in data["items"]:
        size = it[4] if len(it) > 4 else ""
        it[:] = [it[0], it[1], it[2], it[3], size, norm(it[1])]
    JP_LIST = data
    JP_BY_CODE = {it[0]: it for it in data["items"]}
    return data


# よく知られたブランド名・通称 → 証券コード（一覧に存在するものだけ表示）
ALIASES = {
    "ユニクロ": ["9983"], "ジーユー": ["9983"], "セブンイレブン": ["3382"], "マクドナルド": ["2702"], "ニトリ": ["9843"],
    "無印良品": ["7453"], "ドコモ": ["9432"], "エヌティティ": ["9432"], "ntt": ["9432"], "エーユー": ["9433"], "au": ["9433"],
    "ソフトバンク": ["9434", "9984"], "楽天": ["4755"], "メルカリ": ["4385"], "ゾゾ": ["3092"], "しまむら": ["8227"],
    "ワークマン": ["7564"], "吉野家": ["9861"], "すき家": ["7550"], "スシロー": ["3563"], "くら寿司": ["2695"],
    "サイゼリヤ": ["7581"], "丸亀製麺": ["3397"], "コメダ": ["3543"], "ディズニー": ["4661"], "ディズニーランド": ["4661"],
    "キリン": ["2503"], "アサヒ": ["2502"], "サントリー": ["2587"], "ヤクルト": ["2267"], "日清": ["2897"], "カルビー": ["2229"],
    "マツキヨ": ["3088"], "プレステ": ["6758"], "プレイステーション": ["6758"], "スイッチ": ["7974"], "パナソニック": ["6752"],
    "ホンダ": ["7267"], "日産": ["7201"], "スバル": ["7270"], "マツダ": ["7261"], "ヤマト": ["9064"], "クロネコ": ["9064"],
    "佐川": ["9143"], "jal": ["9201"], "ana": ["9202"], "jr東日本": ["9020"], "jr東海": ["9022"], "jr西日本": ["9021"],
    "ユニチャーム": ["8113"], "資生堂": ["4911"], "花王": ["4452"], "ライオン": ["4912"], "キーエンス": ["6861"],
    "東エレ": ["8035"], "レーザーテック": ["6920"], "三菱ufj": ["8306"], "三井住友": ["8316"], "みずほ": ["8411"],
}
ALIASES_N = [(norm(k), v) for k, v in ALIASES.items()]

# 米国株：日本語（カタカナ）の呼び名や、Yahooの検索で出にくい略称 → ティッカー
US_ALIASES = {
    "アップル": ("AAPL", "Apple"), "マイクロソフト": ("MSFT", "Microsoft"), "エヌビディア": ("NVDA", "NVIDIA"), "エヌヴィディア": ("NVDA", "NVIDIA"),
    "アマゾン": ("AMZN", "Amazon.com"), "グーグル": ("GOOGL", "Alphabet（Google）"), "アルファベット": ("GOOGL", "Alphabet（Google）"), "google": ("GOOGL", "Alphabet（Google）"),
    "メタ": ("META", "Meta Platforms（Facebook）"), "フェイスブック": ("META", "Meta Platforms（Facebook）"), "facebook": ("META", "Meta Platforms（Facebook）"),
    "テスラ": ("TSLA", "Tesla"), "ネットフリックス": ("NFLX", "Netflix"), "ブロードコム": ("AVGO", "Broadcom"), "インテル": ("INTC", "Intel"),
    "クアルコム": ("QCOM", "Qualcomm"), "オラクル": ("ORCL", "Oracle"), "セールスフォース": ("CRM", "Salesforce"), "アドビ": ("ADBE", "Adobe"),
    "パランティア": ("PLTR", "Palantir"), "ウーバー": ("UBER", "Uber"), "エアビーアンドビー": ("ABNB", "Airbnb"),
    "バークシャー": ("BRK-B", "Berkshire Hathaway"), "berkshire": ("BRK-B", "Berkshire Hathaway"),
    "コカコーラ": ("KO", "Coca-Cola"), "コカ・コーラ": ("KO", "Coca-Cola"), "coke": ("KO", "Coca-Cola"), "ペプシ": ("PEP", "PepsiCo"),
    "ジョンソンエンドジョンソン": ("JNJ", "Johnson & Johnson"), "ジョンソン・エンド・ジョンソン": ("JNJ", "Johnson & Johnson"), "j&j": ("JNJ", "Johnson & Johnson"),
    "プロクターアンドギャンブル": ("PG", "Procter & Gamble"), "p&g": ("PG", "Procter & Gamble"),
    "マクドナルド": ("MCD", "McDonald's"), "スターバックス": ("SBUX", "Starbucks"), "スタバ": ("SBUX", "Starbucks"), "ナイキ": ("NKE", "Nike"),
    "ディズニー": ("DIS", "Walt Disney"), "コストコ": ("COST", "Costco"), "ウォルマート": ("WMT", "Walmart"), "ホームデポ": ("HD", "Home Depot"),
    "ビザ": ("V", "Visa"), "マスターカード": ("MA", "Mastercard"), "アメックス": ("AXP", "American Express"), "アメリカンエキスプレス": ("AXP", "American Express"),
    "jpモルガン": ("JPM", "JPMorgan Chase"), "ゴールドマン": ("GS", "Goldman Sachs"), "バンクオブアメリカ": ("BAC", "Bank of America"),
    "エクソン": ("XOM", "Exxon Mobil"), "シェブロン": ("CVX", "Chevron"), "ボーイング": ("BA", "Boeing"), "キャタピラー": ("CAT", "Caterpillar"),
    "ファイザー": ("PFE", "Pfizer"), "イーライリリー": ("LLY", "Eli Lilly"), "リリー": ("LLY", "Eli Lilly"), "メルク": ("MRK", "Merck"),
    "アッヴィ": ("ABBV", "AbbVie"), "ユナイテッドヘルス": ("UNH", "UnitedHealth"), "ノボノルディスク": ("NVO", "Novo Nordisk（ADR）"),
    "tsmc": ("TSM", "Taiwan Semiconductor（ADR）"), "台湾積体電路": ("TSM", "Taiwan Semiconductor（ADR）"), "台積電": ("TSM", "Taiwan Semiconductor（ADR）"),
    "asml": ("ASML", "ASML Holding"), "アリババ": ("BABA", "Alibaba（ADR）"), "フィリップモリス": ("PM", "Philip Morris"), "アルトリア": ("MO", "Altria"),
    "ベライゾン": ("VZ", "Verizon"), "at&t": ("T", "AT&T"), "スリーエム": ("MMM", "3M"), "ibm": ("IBM", "IBM"),
}
US_ALIASES_N = [(norm(k), v) for k, v in US_ALIASES.items()]


def alias_hit(an, qn):
    """通称の一致判定（短い英字は完全一致か3文字以上の前方一致のみ）"""
    if not qn or len(qn) < 2:
        return False
    if an == qn:
        return True
    if not re.search(r"[^\x00-\x7F]", an) and len(qn) < 3:
        return False
    if an.startswith(qn):
        return True
    return len(an) >= 3 and qn.startswith(an)

US_EXCHANGES = {"NMS": "NASDAQ", "NGM": "NASDAQ", "NCM": "NASDAQ", "NYQ": "NYSE", "ASE": "NYSE American", "PCX": "NYSE Arca", "BTS": "Cboe"}


def seg_label(seg):
    m = re.match(r"(プライム|スタンダード|グロース)", seg or "")
    return f"東証{m.group(1)}" if m else (seg or "東証")


def jp_item(code, why=""):
    it = JP_BY_CODE.get(code)
    if not it:
        return None
    return {"symbol": f"{code}.T", "code": code, "market": "JP", "name": it[1], "exchange": seg_label(it[2]), "sector": it[3], "why": why}


def search(q):
    q = (q or "").strip()
    if not q:
        return []
    cached = SEARCH_CACHE.get(q)
    if cached is not None:
        return cached
    out, seen = [], set()

    def add(item):
        if item and item["symbol"] not in seen:
            seen.add(item["symbol"])
            out.append(item)

    qu = unicodedata.normalize("NFKC", q).upper()
    qn = norm(q)
    ascii_q = not re.search(r"[^\x00-\x7F]", qu)
    if re.fullmatch(r"\d{3}[0-9A-Z]", qu):
        add(jp_item(qu))
    if re.fullmatch(r"\d{1,3}", qu):
        for it in JP_LIST["items"]:
            if it[0].startswith(qu):
                add(jp_item(it[0]))
            if len(out) >= 8:
                break
    if ascii_q and re.search(r"[A-Z0-9]", qu):
        try:
            res = yf.Search(q, max_results=10, news_count=0, lists_count=0).quotes
        except Exception:
            res = []
        ys = []
        for x in res:
            sym, ex, qt = x.get("symbol") or "", x.get("exchange") or "", x.get("quoteType")
            if qt != "EQUITY":
                continue
            if sym.endswith(".T"):
                ys.append(jp_item(sym[:-2]) or {"symbol": sym, "code": sym[:-2], "market": "JP", "name": x.get("longname") or x.get("shortname") or sym, "exchange": "東証", "sector": "", "why": ""})
            elif ex in US_EXCHANGES:
                ys.append({"symbol": sym, "code": sym, "market": "US", "name": x.get("longname") or x.get("shortname") or sym, "exchange": US_EXCHANGES[ex], "sector": x.get("industryDisp") or x.get("industry") or "", "why": ""})
        ys.sort(key=lambda it: it["code"].upper() != qu)
        for it in ys:
            add(it)
    for an, codes in ALIASES_N:
        if alias_hit(an, qn):
            for c in codes:
                add(jp_item(c, "通称・ブランド名から"))
    for an, (tk, nm) in US_ALIASES_N:
        if alias_hit(an, qn):
            add({"symbol": tk, "code": tk, "market": "US", "name": nm, "exchange": "米国株", "sector": "", "why": "通称・ブランド名から"})
    if qn:
        hits = [it for it in JP_LIST["items"] if (it[5].startswith(qn) if ascii_q else qn in it[5])]
        hits.sort(key=lambda it: (not it[5].startswith(qn), SIZE_RANK.get(it[4], 5), len(it[1]), it[0]))
        for it in hits[:12]:
            add(jp_item(it[0]))
    out = out[:15]
    SEARCH_CACHE.set(q, out)
    return out


# ---------------------------------------------------------------------------
# 業種の対応づけ
# ---------------------------------------------------------------------------
SECTOR_RULES = [
    (("Auto Manufacturers", "Auto Parts", "Auto & Truck"), "auto"),
    (("Steel", "Aluminum", "Copper", "Metal Fabrication", "Other Industrial Metals"), "steel"),
    (("Gold", "Silver", "Industrial Metals & Mining", "Precious Metals", "Coking Coal", "Thermal Coal", "Uranium", "Oil & Gas"), "energy"),
    (("Chemicals", "Agricultural Inputs", "Paper", "Building Materials", "Packaging", "Lumber"), "chemicals"),
    (("Marine Shipping", "Airlines"), "shipping"),
    (("Semiconductor",), "semis"),
    (("Utilities",), "utility"),
    (("Telecom",), "telecom"),
    (("Drug Manufacturers", "Biotechnology", "Medical", "Diagnostics", "Healthcare", "Health Information", "Pharmaceutical"), "pharma"),
    (("Software", "Information Technology Services"), "software"),
    (("Internet Content", "Electronic Gaming", "Entertainment", "Advertising", "Publishing", "Broadcasting"), "internet"),
    (("Banks",), "bank"),
    (("Insurance",), "insurance"),
    (("Capital Markets", "Asset Management", "Credit Services", "Financial", "Mortgage"), "finance"),
    (("REIT", "Real Estate"), "realestate"),
    (("Residential Construction", "Engineering & Construction", "Infrastructure Operations", "Building Products"), "construction"),
    (("Machinery", "Tools & Accessories", "Electrical Equipment", "Aerospace", "Pollution"), "machinery"),
    (("Apparel Manufacturing", "Footwear", "Textile", "Luxury"), "apparel"),
    (("Grocery", "Discount Stores", "Department Stores", "Specialty Retail", "Home Improvement", "Apparel Retail", "Restaurants", "Internet Retail"), "retail"),
    (("Packaged Foods", "Beverages", "Confectioners", "Farm Products", "Food Distribution", "Tobacco", "Household & Personal"), "food"),
    (("Consumer Electronics", "Furnishings", "Recreational Vehicles", "Leisure"), "consumer"),
    (("Electronic Components", "Scientific & Technical", "Computer Hardware", "Communication Equipment", "Electronics & Computer", "Solar"), "electronics"),
    (("Railroads", "Trucking", "Freight", "Airports", "Integrated Freight"), "transport"),
    (("Conglomerates", "Industrial Distribution"), "trading"),
    (("Staffing", "Education", "Business Services", "Consulting", "Rental", "Security", "Personal Services", "Waste", "Travel", "Lodging", "Resorts", "Gambling"), "service"),
]
JPX33 = {
    "水産・農林業": "food", "鉱業": "energy", "建設業": "construction", "食料品": "food", "繊維製品": "apparel", "パルプ・紙": "chemicals",
    "化学": "chemicals", "医薬品": "pharma", "石油・石炭製品": "energy", "ゴム製品": "chemicals", "ガラス・土石製品": "chemicals",
    "鉄鋼": "steel", "非鉄金属": "steel", "金属製品": "machinery", "機械": "machinery", "電気機器": "electronics", "輸送用機器": "auto",
    "精密機器": "electronics", "その他製品": "consumer", "電気・ガス業": "utility", "陸運業": "transport", "海運業": "shipping",
    "空運業": "shipping", "倉庫・運輸関連業": "transport", "情報・通信業": "software", "卸売業": "trading", "小売業": "retail",
    "銀行業": "bank", "証券、商品先物取引業": "finance", "保険業": "insurance", "その他金融業": "finance", "不動産業": "realestate", "サービス業": "service",
}


def map_sector(industry, sector, sector33=""):
    if sector33 == "卸売業":
        return "trading"
    if "Conglomerates" in (industry or "") and not sector33:
        return "machinery"
    text = f"{industry or ''} | {sector or ''}"
    for keys, key in SECTOR_RULES:
        if any(k.lower() in text.lower() for k in keys):
            return key
    return JPX33.get(sector33, "other")


# ---------------------------------------------------------------------------
# 市場（指数）
# ---------------------------------------------------------------------------
INDEX = {"JP": ("1306.T", "TOPIX連動ETF（1306）"), "US": ("^GSPC", "S&P 500")}


def close_series(h):
    c = h["Close"].dropna()
    if getattr(c.index, "tz", None) is not None:
        c.index = c.index.tz_localize(None)
    return c


def ret_1y(c):
    if len(c) < 30:
        return None
    ago = c[c.index <= c.index[-1] - pd.Timedelta(days=365)]
    return (float(c.iloc[-1]) / float(ago.iloc[-1]) - 1) * 100 if len(ago) else None


def market_stats(mk):
    cached = MARKET_CACHE.get(mk)
    if cached:
        return cached
    sym, label = INDEX[mk]
    c = close_series(yf.Ticker(sym).history(period="2y", auto_adjust=False))
    last = float(c.iloc[-1])
    ma50, ma200 = float(c.iloc[-50:].mean()), float(c.iloc[-200:].mean())
    trend = "up" if (last > ma200 and ma50 > ma200) else "down" if (last < ma200 and ma50 < ma200) else "correction"
    out = {"market": mk, "symbol": sym, "label": label, "close": round(last, 2), "ma50": round(ma50, 2), "ma200": round(ma200, 2),
           "ret52": r2(ret_1y(c), 2), "trend": trend, "asOf": c.index[-1].strftime("%Y-%m-%d")}
    MARKET_CACHE.set(mk, out)
    return out


# ---------------------------------------------------------------------------
# 銘柄データ
# ---------------------------------------------------------------------------
def parse_symbol(raw):
    sym = unicodedata.normalize("NFKC", (raw or "").strip()).upper()
    if re.fullmatch(r"\d{3}[0-9A-Z]", sym):
        sym += ".T"
    if not re.fullmatch(r"[A-Z0-9.\-]{1,12}", sym):
        raise ValueError("銘柄コード・ティッカーの形式が正しくありません")
    mk = "JP" if sym.endswith(".T") else "US"
    return sym, mk, (sym[:-2] if mk == "JP" else sym)


FX_CACHE = TTLCache(60 * 60)


def fx_series(src, dst):
    """1 src 通貨あたりの dst 通貨（日次の終値・過去10年）。取得できなければ None"""
    key = f"{src}{dst}"
    hit = FX_CACHE.get(key)
    if hit is not None:
        return hit
    out = None
    for pair, inv in ((f"{src}{dst}=X", False), (f"{dst}{src}=X", True)):
        try:
            h = yf.Ticker(pair).history(period="10y")
            c = close_series(h) if h is not None and not h.empty else None
            if c is not None and len(c):
                out = (1.0 / c) if inv else c
                break
        except Exception:
            continue
    if out is not None:
        FX_CACHE.set(key, out)
    return out


NICE_RATIOS = [0.01, 0.02, 0.025, 0.04, 0.05, 0.1, 0.125, 0.2, 0.25, 1 / 3, 0.5, 1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 14, 15, 16, 20, 25, 30, 40, 50, 100]


def snap_ratio(x):
    """ADRの比率（1ADRあたりの株数）を、よくある値に丸める。はっきりしなければ None"""
    if not x or x <= 0:
        return None
    best = min(NICE_RATIOS, key=lambda r: abs(math.log(x / r)))
    return round(best, 4) if abs(math.log(x / best)) <= math.log(1.2) else None


def hist_pe(close, eps_pairs):
    """過去の平均PER：各決算期の1年間の平均株価 ÷ その期のEPS"""
    pes = []
    for d, e in valid(eps_pairs):
        if e <= 0:
            continue
        seg = close[(close.index > d - pd.Timedelta(days=365)) & (close.index <= d)]
        if len(seg) > 100:
            pe = float(seg.mean()) / e
            if 0 < pe < 200:
                pes.append(pe)
    return round(sum(pes) / len(pes), 1) if len(pes) >= 2 else None


def clean_dividends(hist):
    """配当履歴（株式分割日の未調整の配当を、分割比率で割って直す）"""
    if hist is None or hist.empty or "Dividends" not in hist.columns:
        return None, None
    h = hist.copy()
    if getattr(h.index, "tz", None) is not None:
        h.index = h.index.tz_localize(None)
    div = h["Dividends"]
    first_year = int(h.index[0].year)
    div = div[div > 0].copy()
    if "Stock Splits" in h.columns and not div.empty:
        sp = h["Stock Splits"]
        sp = sp[sp > 1]
        for d in list(div.index):
            for s_date, f in sp.items():
                if abs((d - s_date).days) <= 10:
                    prev = div[div.index < d].tail(4)
                    if len(prev) and div[d] > 2 * float(prev.median()):
                        div[d] = div[d] / float(f)
    return div, first_year


def dividend_streaks(div, first_year, today):
    if div is None:
        return None, None, ""
    last_full = today.year - 1
    if div.empty:
        return 0, 0, ""
    by = div.groupby(div.index.year).sum()
    y, years = last_full, 0
    while y in by.index and by[y] > 0:
        years += 1
        y -= 1
    y, inc = last_full, 0
    while y in by.index and (y - 1) in by.index and by[y] > by[y - 1] * 1.001:
        inc += 1
        y -= 1
    note = f"連続配当・連続増配の年数は、Yahoo Financeの配当履歴（{first_year}年〜）を暦年で集計した推定値です" if years else ""
    if years and last_full - years + 1 <= first_year:
        note += "。データの始まりまで途切れていないため、実際はもっと長い可能性があります"
    return years, inc, note


def fetch_stock(raw, fresh=False):
    sym, mk, code = parse_symbol(raw)
    cached = None if fresh else STOCK_CACHE.get(sym)
    if cached:
        return cached
    with FETCH_SEM:
        t0 = time.time()
        tk = yf.Ticker(sym)
        info = tk.info or {}
        price = num(info.get("currentPrice")) or num(info.get("regularMarketPrice")) or num(info.get("previousClose"))
        if not price:
            raise LookupError("株価を取得できませんでした（コードの誤り、上場廃止、または一時的な取得制限の可能性があります）")
        notes = []
        cur, fcur = info.get("currency"), info.get("financialCurrency")
        conv, fx, fx_hist = False, 1.0, None
        if cur and fcur and cur != fcur:
            fx_hist = fx_series(fcur, cur)
            if fx_hist is not None and len(fx_hist):
                conv, fx = True, float(fx_hist.iloc[-1])
            else:
                notes.append(f"財務データの通貨（{fcur}）が株価の通貨（{cur}）と異なり、為替レートを取得できなかったため換算していません。PERなどが正しく計算できません")
        inc, bs, cf = tk.income_stmt, tk.balance_sheet, tk.cashflow
        try:
            qbs, qcf = tk.quarterly_balance_sheet, tk.quarterly_cashflow
        except Exception:
            qbs, qcf = None, None
        hist = tk.history(period="max", auto_adjust=False)
        close = close_series(hist) if hist is not None and not hist.empty else None
        today = dt.date.today()

        eps_p = series(inc, "Diluted EPS", "Basic EPS")
        eps_raw = dict(valid(eps_p))
        disc_p = dict(valid(series(inc, "Net Income Discontinuous Operations")))
        ni_p = dict(valid(series(inc, "Net Income Common Stockholders", "Net Income")))
        avg_sh = dict(valid(series(inc, "Diluted Average Shares", "Basic Average Shares")))
        adj_years = []

        def eps_basis(d, e):
            """そのEPSが何株あたりか（純利益÷EPS）。ADRではADR単位のことがある"""
            ni = ni_p.get(d)
            if ni and e and abs(e) > 1e-9 and ni / e > 0:
                return ni / e
            return avg_sh.get(d)

        def ex_disc(d, e):
            # 非継続事業（売却・分離した事業）の損益を除いた、継続事業ベースのEPS
            dv = disc_p.get(d)
            if e is None or not dv:
                return e
            sh = eps_basis(d, e)
            if not sh:
                return e
            per = dv / sh
            if abs(per) > max(abs(e) * 0.05, 1e-9):
                adj_years.append(d.year)
                return e - per
            return e

        eps_p = [(d, ex_disc(d, e)) for d, e in eps_p]
        if adj_years:
            notes.append("事業の売却・分離などによる「非継続事業」の損益を除いたEPS（継続事業ベース）を使いました（"
                         + "・".join(f"{y}年" for y in sorted(set(adj_years))) + "の決算期）")
        eps_v = sorted(valid(eps_p))[-5:]
        eq_p = dict(valid(series(bs, "Stockholders Equity", "Common Stock Equity")))
        sh_p = dict(valid(series(bs, "Ordinary Shares Number", "Share Issued")))

        def bps_at(d):
            for dd, e in eq_p.items():
                if abs((dd - d).days) <= 40:
                    s = next((v for k, v in sh_p.items() if abs((k - dd).days) <= 40 and v), None)
                    if s:
                        return round(e / s, 3)
            return ""

        eps_hist = [r2(v) for _, v in eps_v]
        bps_hist = [bps_at(d) for d, _ in eps_v]

        sales = valid(series(inc, "Total Revenue", "Operating Revenue"))
        ebit = valid(series(inc, "Operating Income", "EBIT"))
        gross = first(series(inc, "Gross Profit"))
        pretax = first(series(inc, "Pretax Income"))
        rnd = first(series(inc, "Research And Development"))
        intexp = first(series(inc, "Interest Expense", "Interest Expense Non Operating"))
        inv = valid(series(bs, "Inventory"))
        if not inv and bs is not None and not bs.empty:
            inv = [(None, 0.0), (None, 0.0)]

        def lvl(*names):
            v = first(series(qbs, *names))
            return v if v is not None else first(series(bs, *names))

        equity = lvl("Stockholders Equity", "Common Stock Equity")
        cur_a, cur_l = lvl("Current Assets"), lvl("Current Liabilities")
        tot_l, tot_a = lvl("Total Liabilities Net Minority Interest"), lvl("Total Assets")
        ppe = lvl("Net PPE")
        lt_debt = lvl("Long Term Debt", "Long Term Debt And Capital Lease Obligation")

        capex_q = [v for _, v in series(qcf, "Capital Expenditure")[:4] if v is not None]
        capex = abs(sum(capex_q)) if len(capex_q) == 4 else (abs(first(series(cf, "Capital Expenditure"))) if first(series(cf, "Capital Expenditure")) is not None else None)
        opcf = num(info.get("operatingCashflow"))
        if opcf is None:
            opcf = first(series(cf, "Operating Cash Flow"))

        mcap_info = num(info.get("marketCap"))
        shares = (mcap_info / price) if mcap_info else (num(info.get("sharesOutstanding")) or num(info.get("impliedSharesOutstanding")))
        cash_v, debt_v = num(info.get("totalCash")), num(info.get("totalDebt"))
        pe_pairs = eps_p
        if conv:
            # 外国企業（ADRなど）：財務の通貨を株価の通貨に換算し、1株あたりの値をADR単位にそろえる
            def cv(v):
                return None if v is None else v * fx

            def fx_avg(d):
                seg = fx_hist[(fx_hist.index > d - pd.Timedelta(days=365)) & (fx_hist.index <= d)]
                return float(seg.mean()) if len(seg) > 20 else fx

            r_eps = None
            if eps_v and shares:
                d0 = eps_v[-1][0]
                b0 = eps_basis(d0, eps_raw.get(d0))
                r_eps = snap_ratio(b0 / shares) if b0 else None
            ord_now = first(series(bs, "Ordinary Shares Number", "Share Issued"))
            r_bps = (ord_now / shares) if (ord_now and shares) else r_eps
            if r_eps is None:
                eps_hist, pe_pairs = [], []
                notes.append("ADR（預託証券）と元の株式の比率が分からないため、EPSの推移は空欄にしました")
            else:
                eps_hist = [r2(v * r_eps * fx) for v in eps_hist]
                pe_pairs = [(d, v * r_eps * fx_avg(d)) for d, v in valid(eps_p)]
            bps_hist = [round(v * r_bps * fx, 3) if (v not in ("", None) and r_bps) else "" for v in bps_hist]
            equity, cur_a, cur_l, tot_l, tot_a, ppe, lt_debt = map(cv, (equity, cur_a, cur_l, tot_l, tot_a, ppe, lt_debt))
            capex, opcf, gross, pretax, rnd, intexp, cash_v, debt_v = map(cv, (capex, opcf, gross, pretax, rnd, intexp, cash_v, debt_v))
            sales = [(d, v * fx) for d, v in sales]
            ebit = [(d, v * fx) for d, v in ebit]
            inv = [(d, v * fx) for d, v in inv]
            notes.append(f"財務データ（{fcur}）を現在の為替レート（1{fcur}＝{fx:.4g}{cur}）で{cur}に換算しました"
                         + (f"。1株あたりの値はADR単位（1ADR＝{r_eps:g}株）にそろえています" if r_eps and r_eps != 1 else ""))
            if fcur == "JPY":
                notes.append("日本企業の米国預託証券（ADR）です。東証の銘柄コードで検索して分析すると、円ベースの正確なデータを使えます")
        sh_hist = [v for _, v in sorted(sh_p.items())]
        shares_chg = round((sh_hist[-1] / sh_hist[0] - 1) * 100, 2) if len(sh_hist) >= 2 and sh_hist[0] else None
        rep = first(series(cf, "Repurchase Of Capital Stock"))
        if conv and rep is not None:
            rep *= fx
        mcap = price * shares if shares else None
        if rep is not None:
            buyback = "yes" if (rep < 0 and (not mcap or abs(rep) > mcap * 0.001)) else "no"
        elif shares_chg is not None and shares_chg <= -1:
            buyback = "yes"
        else:
            buyback = "unknown"

        dps = num(info.get("dividendRate"))
        div, first_year = clean_dividends(hist)
        div_years, div_inc, div_note = dividend_streaks(div, first_year, today)
        ttm_div = None
        if div is not None and close is not None:
            ttm_div = float(div[div.index > close.index[-1] - pd.Timedelta(days=365)].sum())
        if dps and ttm_div and not (0.4 < dps / ttm_div < 2.5):
            notes.append(f"Yahoo Financeの予想配当（{dps:g}）が直近1年の配当実績と大きく違うため、実績（{ttm_div:g}）を使いました（株式分割の反映ずれなどの可能性）")
            dps = round(ttm_div, 4)
        if dps is None:
            dps = num(info.get("trailingAnnualDividendRate"))
        if dps is None and ttm_div is not None:
            dps = round(ttm_div, 4)
        if dps is None and div_years == 0:
            dps = 0.0

        pe_avg = hist_pe(close, pe_pairs) if close is not None else None
        ret52 = ret_1y(close) if close is not None else None
        high52 = num(info.get("fiftyTwoWeekHigh"))
        if high52 is None and close is not None:
            high52 = float(close[close.index > close.index[-1] - pd.Timedelta(days=365)].max())
        try:
            idx52 = market_stats(mk)["ret52"]
        except Exception:
            idx52 = None

        t_eps = num(info.get("trailingEps"))
        if t_eps and eps_hist and eps_hist[-1] and t_eps * eps_hist[-1] > 0:
            ratio = t_eps / eps_hist[-1]
            if ratio > 3 or ratio < 1 / 3:
                notes.append("直近のEPSと過去のEPSの水準が大きく違います（株式分割の反映ずれや一時的な損益の可能性）。EPSの推移を決算資料で確認してください")
        elif t_eps and eps_hist and eps_hist[-1] and t_eps * eps_hist[-1] < 0:
            notes.append(f"直近12か月のEPS（{t_eps:g}）と直近の年度のEPS（{eps_hist[-1]:g}）で黒字・赤字が逆です（一時的な損益の可能性）。決算資料で確認してください")
        f_eps = num(info.get("forwardEps"))
        if f_eps and t_eps and f_eps * t_eps > 0 and not (1 / 3 < f_eps / t_eps < 3):
            fr = f_eps / t_eps
            split_f = None
            if hist is not None and not hist.empty and "Stock Splits" in hist.columns:
                sp = hist["Stock Splits"]
                sp = sp[sp > 1]
                if len(sp) and (hist.index[-1] - sp.index[-1]).days <= 400:
                    split_f = float(sp.iloc[-1])
            if split_f and 0.6 < fr / split_f < 1.6:
                f_eps = f_eps / split_f
                notes.append(f"予想EPSが株式分割（1→{split_f:g}株）を反映していなかったため、分割後の基準に換算しました")
            else:
                notes.append("予想EPSが実績EPSと大きく違うため、使わずに空欄にしました（必要なら会社予想を手で入力してください）")
                f_eps = None

        sh_now = shares
        bps_now = (equity / sh_now) if (equity and sh_now) else None
        last_bps = next((v for v in reversed(bps_hist) if v not in ("", None)), None)
        if bps_now and last_bps and bps_now > 0 and last_bps > 0 and not (0.5 < bps_now / last_bps < 2):
            f = bps_now / last_bps
            bps_hist = [round(v * f, 3) if v not in ("", None) else "" for v in bps_hist]
            notes.append("BPSの推移を、現在の株価と同じ1株の基準に換算しました（複数の株式クラスがある会社など）")

        jp = JP_BY_CODE.get(code) if mk == "JP" else None
        name = jp[1] if jp else (info.get("longName") or info.get("shortName") or code)
        q_eps = num(info.get("earningsQuarterlyGrowth"))
        q_rev = num(info.get("revenueGrowth"))
        inst = num(info.get("heldPercentInstitutions"))
        rec = {
            "market": mk, "code": code, "name": name, "nameEn": info.get("longName") or info.get("shortName") or "", "yahooSymbol": sym,
            "sector": map_sector(info.get("industry"), info.get("sector"), jp[3] if jp else ""),
            "industryEn": info.get("industry") or "", "sector33": jp[3] if jp else "", "segment": seg_label(jp[2]) if jp else (US_EXCHANGES.get(info.get("exchange"), info.get("exchange") or "")),
            "price": r2(price), "shares": mil(shares), "asOf": (close.index[-1].strftime("%Y-%m-%d") if close is not None and len(close) else today.strftime("%Y-%m-%d")),
            "eps": r2(t_eps) if t_eps is not None else (eps_hist[-1] if eps_hist else None), "epsFwd": r2(f_eps),
            "epsHist": eps_hist, "bpsHist": bps_hist,
            "qEpsG": round(q_eps * 100, 1) if q_eps is not None else None, "qSalesG": round(q_rev * 100, 1) if q_rev is not None else None,
            "dps": r2(dps), "divYears": div_years, "divIncYears": div_inc, "buyback": buyback, "sharesChg": shares_chg,
            "cash": mil(cash_v), "debt": mil(debt_v), "ltDebt": mil(lt_debt), "equity": mil(equity),
            "curAssets": mil(cur_a), "curLiab": mil(cur_l), "totalLiab": mil(tot_l), "totalAssets": mil(tot_a), "ppe": mil(ppe),
            "opCF": mil(opcf), "capex": mil(capex),
            "sales": mil(sales[0][1]) if sales else None, "salesPrev": mil(sales[1][1]) if len(sales) > 1 else None,
            "grossProfit": mil(gross), "ebit": mil(ebit[0][1]) if ebit else None, "ebitPrev": mil(ebit[1][1]) if len(ebit) > 1 else None,
            "pretax": mil(pretax), "rnd": mil(rnd), "interestExp": mil(abs(intexp)) if intexp is not None else None,
            "inventory": mil(inv[0][1]) if inv else None, "inventoryPrev": mil(inv[1][1]) if len(inv) > 1 else None,
            "high52": r2(high52), "ret52": r2(ret52, 2), "idx52": idx52,
            "instPct": round(inst * 100, 1) if inst is not None else None, "analysts": num(info.get("numberOfAnalystOpinions")),
            "peHistAvg": pe_avg, "currency": cur or "",
            "source": "yahoo", "fetchedAt": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        }
        if notes:
            rec["fetchNote"] = " / ".join(notes)
        if div_note:
            rec["divNote"] = div_note
        print(f"[取得] {sym:<9} {name[:24]}（{time.time() - t0:.1f}秒）")
        STOCK_CACHE.set(sym, rec)
        return rec


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------
class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(HERE), **kw)

    def log_message(self, fmt, *args):
        pass

    def send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        path = urllib.parse.unquote(u.path)
        qs = urllib.parse.parse_qs(u.query)
        if path in ("/", "/index.html"):
            self.send_response(302)
            self.send_header("Location", "/" + urllib.parse.quote(APP_HTML))
            self.end_headers()
            return
        if path.startswith("/api/"):
            return self.api(path, qs)
        if path == "/" + APP_HTML:
            data = (HERE / APP_HTML).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(404)

    def api(self, path, qs):
        q = lambda k: (qs.get(k) or [""])[0]
        try:
            if path == "/api/ping":
                return self.send_json({"ok": True, "app": "pro-investor-analyzer", "version": VERSION, "jpList": JP_LIST.get("date", ""), "jpCount": len(JP_LIST.get("items", []))})
            if path == "/api/search":
                return self.send_json({"results": search(q("q"))})
            if path == "/api/stock":
                return self.send_json({"stock": fetch_stock(q("symbol"), fresh=q("fresh") == "1")})
            if path == "/api/market":
                out = {}
                for mk in ("JP", "US"):
                    try:
                        out[mk] = market_stats(mk)
                    except Exception as e:  # noqa: BLE001
                        out[mk] = {"error": str(e)}
                return self.send_json(out)
            return self.send_json({"error": "不明なAPIです"}, 404)
        except ValueError as e:
            return self.send_json({"error": str(e)}, 400)
        except LookupError as e:
            return self.send_json({"error": str(e)}, 404)
        except Exception as e:  # noqa: BLE001
            msg = str(e)
            if "Too Many Requests" in msg or "Rate limit" in msg or "YFRateLimit" in type(e).__name__:
                return self.send_json({"error": "Yahoo Financeの取得制限に達しました。数分おいてからもう一度お試しください"}, 429)
            print(f"[エラー] {path} {q('symbol') or q('q')}: {type(e).__name__}: {msg[:200]}")
            return self.send_json({"error": f"データを取得できませんでした（{type(e).__name__}）"}, 500)


def already_running(port):
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=2) as r:
            return json.loads(r.read().decode("utf-8")).get("app") == "pro-investor-analyzer"
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser(description="プロ投資家の視点 日米株アナライザー（ローカルサーバー）")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--update-list", action="store_true", help="東証上場銘柄一覧（JPX）を再取得して終了")
    a = ap.parse_args()
    if a.update_list:
        d = load_jp_list(force=True)
        print(f"銘柄一覧を更新しました（{d.get('date', '')}・{len(d['items'])}銘柄）")
        return 0
    url = f"http://127.0.0.1:{a.port}/" + urllib.parse.quote(APP_HTML)
    if already_running(a.port):
        print("すでに起動しています。ブラウザで開きます。")
        if not a.no_browser:
            webbrowser.open(url)
        return 0
    load_jp_list()
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", a.port), Handler)
    except OSError as e:
        print(f"ポート{a.port}を使えませんでした（{e}）。他のソフトが使っている可能性があります。")
        input("Enterキーで終了します…")
        return 1
    print("=" * 64)
    print(" プロ投資家の視点 日米株アナライザー")
    print(f" アプリのアドレス：{url}")
    print(f" 東証上場銘柄一覧：{JP_LIST.get('date', '') or '未取得'}（{len(JP_LIST.get('items', []))}銘柄）")
    print(" 使い終わったら、この画面を閉じてください。")
    print("=" * 64)
    if not a.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
