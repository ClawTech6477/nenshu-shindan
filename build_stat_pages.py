#!/usr/bin/env python3
"""令和7年(2025)賃金構造基本統計調査 から 職種別 / 都道府県別 の静的ページを生成する。

データ出典(e-Stat ファイル。APIキー不要):
  https://www.e-stat.go.jp/stat-search/files?page=1&toukei=00450091&tstat=000001011429
  - 職種 第1表 (statInfId 000040421116) 職種(小分類)・性別 きまって支給する現金給与額 他
  - 職種 第5表 (statInfId 000040421120) 職種(小分類)・年齢階級別
  - 都道府県別 参考表1 (statInfId 000040421202) 性・都道府県別
  - 都道府県別 第1表 (statInfId 000040421167-190) 都道府県・産業・年齢階級別
  - 都道府県別 第3表 (statInfId 000040421195-198) 都道府県・職種(特掲)・性別 → 職種×都道府県ページ
職種×都道府県ページは段階公開: 公開件数は ~/openclaw/data/seo/nenshu_pref_occ.json の rollout
(scripts/seo/nenshu_pref_occ_loop.py が掲載率を見て週1で増減)。労働者数の多い組み合わせから出す。
年収 = きまって支給する現金給与額(月額) x 12 + 年間賞与その他特別給与額。単位は千円→万円に換算。

使い方:
  ~/openclaw/venv/bin/python3 build_stat_pages.py            # ページ生成 + sitemap + 内部リンク
  ~/openclaw/venv/bin/python3 build_stat_pages.py --no-link  # 内部リンク挿入をしない
何度実行しても同じ結果になる(冪等)。
"""
import html
import json
import os
import re
import sys
import unicodedata
import urllib.request
from datetime import date
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent
CACHE = Path.home() / "openclaw" / "data" / "estat_cache"
BASE = "https://nenshu.clawtechjp.com"
YEAR = 2025
REIWA = 7
TODAY = date.today().isoformat()
JP_TODAY = "%d年%d月%d日" % (date.today().year, date.today().month, date.today().day)
GA_ID = "G-RZ1TP2K3N0"
CSS_URL = "https://clawtechjp.com/ui/ct-ui.css"
SRC_LIST_URL = "https://www.e-stat.go.jp/stat-search/files?page=1&toukei=00450091&tstat=000001011429"
SRC_MHLW_URL = "https://www.mhlw.go.jp/toukei/itiran/roudou/chingin/kouzou/z2025/index.html"
SRC_TEXT = "出典: 厚生労働省 賃金構造基本統計調査 令和7年(2025年)"
DL = "https://www.e-stat.go.jp/stat-search/file-download?statInfId=%012d&fileKind=4"

OCC1_ID = 40421116
OCC5_ID = 40421120
PREF_REF1_ID = 40421202
PREF1_IDS = list(range(40421167, 40421191))
PREF3_IDS = [40421195, 40421196, 40421197, 40421198]  # 都道府県別 第3表 職種(特掲)×性別
PO_CONF = Path.home() / "openclaw" / "data" / "seo" / "nenshu_pref_occ.json"
PO_BUILT = Path.home() / "openclaw" / "data" / "seo" / "nenshu_pref_occ_built.json"

# 掲載の足切り(薄いデータは出さない)
MIN_OCC_WORKERS = 5000      # 職種: 企業規模計・男女計の推計労働者数(人)
MIN_SEX_WORKERS = 1000      # 男女別の行を出す最低人数
MIN_AGE_WORKERS = 500       # 年齢階級行を出す最低人数
MIN_SIZE_WORKERS = 1000     # 企業規模別の行を出す最低人数
MIN_PREF_AGE_WORKERS = 2000
MIN_PREF_IND_WORKERS = 5000
MIN_PO_WORKERS = 1000       # 職種×都道府県ページを作る最低人数(男女計)
MIN_PO_TABLE_WORKERS = 300  # 比較表に載せる最低人数
MIN_PO_SEX_WORKERS = 500    # 職種×都道府県の男女別を出す最低人数

PREFS = ["北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "茨城県", "栃木県", "群馬県",
         "埼玉県", "千葉県", "東京都", "神奈川県", "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
         "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
         "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県",
         "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]


def short_pref(name):
    return name if name == "北海道" else name[:-1]


SHORT2CODE = {short_pref(n): i + 1 for i, n in enumerate(PREFS)}

# 職種名キーワード -> affiliates.json の職種キー
JOB_KEYWORDS = [
    ("pharma", ["薬剤師"]),
    ("care", ["介護職員", "訪問介護", "ホームヘルパー", "介護支援専門員", "福祉施設指導専門員"]),
    ("nurse", ["看護師", "准看護師"]),
    ("childcare", ["保育士"]),
    ("it", ["ソフトウェア", "システムコンサルタント", "情報処理", "通信ネットワーク", "プログラマ"]),
    ("constr", ["建築技術者", "土木技術者", "測量技術者", "建設", "大工", "電気工事", "配管", "とび", "左官", "建築"]),
    ("acct", ["経理", "会計", "総務", "人事", "一般事務", "事務従事者", "営業事務", "財務", "法務"]),
]

# ---------------------------------------------------------------- 取得

def fetch(sid):
    CACHE.mkdir(parents=True, exist_ok=True)
    p = CACHE / ("%012d.xlsx" % sid)
    if p.exists() and p.stat().st_size > 20000:
        return p
    req = urllib.request.Request(DL % sid, headers={"User-Agent": "Mozilla/5.0 (nenshu-shindan stat builder)"})
    with urllib.request.urlopen(req, timeout=120) as r:
        data = r.read()
    if data[:2] != b"PK":
        raise RuntimeError("not an xlsx: statInfId=%d (%d bytes)" % (sid, len(data)))
    p.write_bytes(data)
    return p


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").strip())
        except ValueError:
            return None
    return None


def norm(s):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s)))


def rec(row, base):
    """base 列から age,tenure,hours,ot,cash,sched,bonus,workers(十人)"""
    vals = [num(row[base + i]) if base + i < len(row) else None for i in range(8)]
    d = dict(age=vals[0], ten=vals[1], hrs=vals[2], ot=vals[3], cash=vals[4], sch=vals[5], bonus=vals[6],
             wk=(vals[7] * 10 if vals[7] is not None else None))
    if d["cash"] is None or d["bonus"] is None:
        d["annual"] = None
    else:
        d["annual"] = d["cash"] * 12 + d["bonus"]  # 千円
    return d


SIZES = [("all", 3), ("l1000", 11), ("m100", 19), ("s10", 27)]


def parse_occ1():
    wb = openpyxl.load_workbook(fetch(OCC1_ID), read_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    occ, order, sex = {}, [], "男女計"
    for r in rows[11:]:
        if len(r) < 11 or r[1] is None or r[3] is None:
            continue
        parts = [x.strip() for x in str(r[1]).split("\n")]
        parts = [x for x in parts if x]
        if len(parts) > 1 and norm(parts[0]) in ("男女計", "男", "女"):
            sex = norm(parts[0])
        name = norm(parts[-1])
        if name in ("男女計", "男", "女"):
            continue
        if name not in occ:
            occ[name] = {}
            order.append(name)
        occ[name][sex] = {k: rec(r, b) for k, b in SIZES}
    return occ, order


def parse_occ5():
    wb = openpyxl.load_workbook(fetch(OCC5_ID), read_only=True)
    ws = wb["男女計"]
    ages, cur = {}, None
    for r in ws.iter_rows(min_row=12, values_only=True):
        if len(r) < 11 or r[1] is None:
            continue
        label = str(r[1])
        if "歳" in label:
            if cur is None:
                continue
            lab = norm(label)
            d = rec(r, 3)
            if d["annual"] is not None:
                ages[cur].append((lab, d))
        else:
            if r[3] is None:
                continue
            cur = norm(label.split("\n")[-1])
            ages[cur] = []
    return ages


def parse_pref_ref():
    wb = openpyxl.load_workbook(fetch(PREF_REF1_ID), read_only=True)
    out = {}
    for sheet, sexes in (("男女計", ["男女計"]), ("男女別", ["男", "女"])):
        for r in wb[sheet].iter_rows(values_only=True):
            cells = [c for c in r if c is not None]
            if len(cells) < 9 or not isinstance(cells[0], str):
                continue
            nm = norm(cells[0])
            if nm == "全国":
                key = "全国"
            elif nm in SHORT2CODE:
                key = nm
            else:
                continue
            vals = cells[1:]
            for si, sx in enumerate(sexes):
                chunk = vals[si * 8:(si + 1) * 8]
                if len(chunk) < 8:
                    continue
                d = rec([None] * 0 + chunk, 0)
                out.setdefault(key, {})[sx] = d
    return out


def parse_pref_detail():
    ages, inds = {}, {}
    for sid in PREF1_IDS:
        wb = openpyxl.load_workbook(fetch(sid), read_only=True)
        for sn in wb.sheetnames:
            m = re.match(r"^\((.+?)\)(.+)$", sn)
            if not m:
                continue
            pf, ind = m.group(1), m.group(2)
            ind_n = norm(ind)
            if pf not in SHORT2CODE:
                continue
            is_total = ind_n == "産業計"
            is_large = bool(re.match(r"^[A-Z][^0-9]", ind_n))
            if not (is_total or is_large):
                continue
            rows = list(wb[sn].iter_rows(values_only=True))
            tot, agel = None, []
            for r in rows[10:]:
                if len(r) < 11 or r[2] is None:
                    continue
                lab = norm(r[2])
                if lab in ("男", "女"):
                    break  # 以降は男女別ブロック。男女計のみ使う
                if lab == "男女計" and tot is None:
                    tot = rec(r, 3)
                elif "歳" in lab and is_total:
                    d = rec(r, 3)
                    if d["annual"] is not None:
                        agel.append((lab, d))
            if tot is None:
                continue
            if is_total:
                ages[pf] = agel
            else:
                inds.setdefault(pf, []).append((ind_n, tot))
    return ages, inds


def parse_pref_occ():
    """都道府県別 第3表 -> out[県番号(0=全国)][職種名][男女計|男|女] = rec"""
    out = {}
    for sid in PREF3_IDS:
        wb = openpyxl.load_workbook(fetch(sid), read_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
        cols = {}
        for i, v in enumerate(rows[5]):
            if v is None:
                continue
            lab = norm(v)
            m = re.match(r"^(\d{2})", lab)
            if lab == "全国":
                cols[i] = 0
            elif m:
                cols[i] = int(m.group(1))
        sex = "男女計"
        for r in rows[9:]:
            if len(r) < 4 or r[1] is None:
                continue
            parts = [norm(x) for x in str(r[1]).split("\n") if x.strip()]
            if not parts:
                continue
            if len(parts) > 1 and parts[0] in ("男女計", "男", "女"):
                sex = parts[0]
            name = parts[-1]
            if name in ("男女計", "男", "女"):
                continue
            for c, code in cols.items():
                d = rec(r, c)
                if d["annual"] is None or not d["wk"]:
                    continue
                out.setdefault(code, {}).setdefault(name, {})[sex] = d
    return out


# ---------------------------------------------------------------- 表示ヘルパ

def man(x):
    """千円(年額) -> 万円 整数文字列"""
    return "{:,}".format(int(round(x / 10.0)))


def man1(x):
    """千円(月額) -> 万円 小数1桁"""
    return "{:.1f}".format(x / 10.0)


def esc(s):
    return html.escape(str(s), quote=True)


def age_label(lab):
    lab = lab.replace("～", "〜").replace("~", "〜")
    if lab.startswith("〜"):
        return lab[1:] + "以下"
    if lab.endswith("歳〜"):
        return lab[:-2] + "歳以上"
    return lab


def load_affiliates():
    p = REPO / "blog" / "affiliates.json"
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("jobs", {})
    except Exception:
        return {}


def pick_affiliates(name, jobs):
    """職種名に合う承認済みサービス(実効EPC順)。無ければ全体上位2件。"""
    for key, kws in JOB_KEYWORDS:
        if any(k in name for k in kws) and jobs.get(key):
            lst = [s for s in jobs[key] if not s.get("dead")]
            if lst:
                return lst[:2], True
    return [], False


def top_affiliates(jobs, n=2):
    seen, flat = set(), []
    for lst in jobs.values():
        for s in lst:
            if s.get("dead") or s["key"] in seen:
                continue
            seen.add(s["key"])
            flat.append(s)
    flat.sort(key=lambda s: -(s.get("effective_epc") or 0))
    return flat[:n]


def aff_html(page_id, items, matched):
    if not items:
        return ""
    out = []
    head = "この職種の転職を考えるなら" if matched else "職種特化の転職相談先の例"
    for i, s in enumerate(items):
        out.append(
            '<aside class="ct-panel ct-panel--raised ct-cta-card" aria-label="転職サービス">'
            '<span class="ct-pr">PR</span><h2 class="ct-h3">%s: %s</h2><p>%s</p>'
            '<a class="ct-btn ct-btn--primary" href="%s" rel="nofollow sponsored noopener" target="_blank" '
            "onclick=\"if(window.gtag)gtag('event','site_click',{site_name:'%s',rank:%d,placement:'stat_page',page:'%s'})\">"
            '公式サイトで詳しく見る</a></aside>'
            % (esc(head), esc(s["name"]), esc(s.get("desc", "")), esc(s["url"]),
               esc(s["name"]).replace("'", ""), i + 1, esc(page_id)))
    return "\n".join(out)


def table(headers, rows, num_cols=(), first_wrap=False):
    h = "".join('<th%s>%s</th>' % (' class="is-num"' if i in num_cols else "", esc(x)) for i, x in enumerate(headers))
    body = []
    for r in rows:
        tds = []
        for i, c in enumerate(r):
            cls = []
            if i in num_cols:
                cls.append("is-num")
            if i == 0 and first_wrap:
                cls.append("n")
            tds.append('<td%s>%s</td>' % (' class="%s"' % " ".join(cls) if cls else "", c))
        body.append("<tr>%s</tr>" % "".join(tds))
    return ('<div class="ct-scroll-x"><table class="ct-table"><thead><tr>%s</tr></thead><tbody>%s</tbody></table></div>'
            % (h, "".join(body)))


def page_shell(path, title, desc, crumbs, body_html, jobj_extra=None):
    url = BASE + path
    crumb_ld = {
        "@context": "https://schema.org", "@type": "BreadcrumbList",
        "itemListElement": [
            {"@type": "ListItem", "position": i + 1, "name": n, **({"item": BASE + u} if u else {})}
            for i, (n, u) in enumerate(crumbs)],
    }
    crumb_html = " › ".join(
        '<a href="%s">%s</a>' % (esc(u), esc(n)) if u and i < len(crumbs) - 1 else esc(n)
        for i, (n, u) in enumerate(crumbs))
    return """<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>%(title)s</title>
<meta name="description" content="%(desc)s">
<link rel="canonical" href="%(url)s">
<meta property="og:type" content="article">
<meta property="og:title" content="%(title)s">
<meta property="og:description" content="%(desc)s">
<meta property="og:url" content="%(url)s">
<meta property="og:image" content="%(base)s/og.png">
<meta property="og:site_name" content="年収UPシミュレーター">
<meta property="og:locale" content="ja_JP">
<meta name="twitter:card" content="summary">
<script type="application/ld+json">%(ld)s</script>%(ld2)s
<link rel="stylesheet" href="%(css)s">
<style>.ct-table td.n{white-space:normal;min-width:11em}.ct-body{max-width:none}.ct-body p{max-width:46em}html,body{overflow-x:hidden}</style>
<script async src="https://www.googletagmanager.com/gtag/js?id=%(ga)s"></script>
<script>window.dataLayer=window.dataLayer||[];function gtag(){dataLayer.push(arguments);}gtag('js',new Date());gtag('config','%(ga)s');</script>
</head>
<body>
<header class="ct-header"><div class="ct-container ct-header__in"><a href="/" style="text-decoration:none;color:inherit;font-weight:700">年収UPシミュレーター</a></div></header>
<nav class="ct-container ct-crumb" aria-label="パンくず">%(crumb)s</nav>
<main class="ct-container">
%(body)s
</main>
</body></html>
""" % dict(title=esc(title), desc=esc(desc), url=url, base=BASE, css=CSS_URL, ga=GA_ID,
           ld=json.dumps(crumb_ld, ensure_ascii=False), crumb=crumb_html, body=body_html,
           ld2=('\n<script type="application/ld+json">%s</script>' % json.dumps(jobj_extra, ensure_ascii=False)) if jobj_extra else "")


def cta_diagnosis(page_id):
    return ('<aside class="ct-panel ct-panel--raised ct-cta-card" aria-label="診断">'
            '<p class="ct-eyebrow">4つの質問・30秒</p><h2 class="ct-h3">あなたの適正年収を診断する</h2>'
            '<p>平均はあくまで全体の目安です。業種・年代・今後の志向を選ぶと、あなたの適正年収と上げられる金額の目安がわかります。登録不要・無料です。</p>'
            '<a class="ct-btn ct-btn--primary" href="/" onclick="if(window.gtag)gtag(\'event\',\'internal_click\','
            "{src:'%s',to:'diagnosis'})\">無料で診断する</a></aside>" % esc(page_id))


def source_note():
    return ('<p class="ct-note">%s。<a href="%s" rel="noopener" target="_blank">e-Stat 統計表ファイル</a>・'
            '<a href="%s" rel="noopener" target="_blank">結果の概況</a>を当サイトで加工して作成。'
            '年収は「きまって支給する現金給与額(月額)×12+年間賞与その他特別給与額」で算出した目安で、'
            '企業規模10人以上の事業所の一般労働者(短時間労働者を除く)の平均です。'
            '推計値のため実際の年収を保証するものではありません。最終確認日: %s</p>'
            % (SRC_TEXT, SRC_LIST_URL, SRC_MHLW_URL, JP_TODAY))


# ---------------------------------------------------------------- ページ生成

def build(argv):
    do_link = "--no-link" not in argv
    occ, order = parse_occ1()
    ages5 = parse_occ5()
    pref_ref = parse_pref_ref()
    pref_ages, pref_inds = parse_pref_detail()
    jobs = load_affiliates()
    nat = pref_ref["全国"]["男女計"]
    nat_annual = nat["annual"]

    # --- 職種の選別
    slug = {n: "j%03d" % (i + 1) for i, n in enumerate(order)}
    keep = []
    for n in order:
        a = occ[n].get("男女計", {}).get("all")
        if not a or a["annual"] is None or a["wk"] is None or a["wk"] < MIN_OCC_WORKERS:
            continue
        keep.append(n)
    keep.sort(key=lambda n: -occ[n]["男女計"]["all"]["annual"])
    rank = {n: i + 1 for i, n in enumerate(keep)}
    N = len(keep)

    # --- 職種×都道府県(段階公開)
    po = parse_pref_occ()
    try:
        rollout = int(json.loads(PO_CONF.read_text())["rollout"])
    except Exception:
        rollout = 1000
    po_cands = [(c, n) for c in range(1, 48) for n in keep
                if (po.get(c, {}).get(n, {}).get("男女計") or {}).get("wk", 0) >= MIN_PO_WORKERS]
    po_cands.sort(key=lambda t: -po[t[0]][t[1]]["男女計"]["wk"])
    po_live = set(po_cands[:rollout])

    def po_path(c, n):
        return "/nenshu/pref/%02d/%s/" % (c, slug[n])

    def po_link(c, n, label):
        return ('<a href="%s">%s</a>' % (po_path(c, n), esc(label))) if (c, n) in po_live else esc(label)

    def po_rows_for(n):  # その職種の都道府県別(比較表に載せられるもの)を年収順
        rs = [(c, po[c][n]["男女計"]) for c in range(1, 48)
              if (po.get(c, {}).get(n, {}).get("男女計") or {}).get("wk", 0) >= MIN_PO_TABLE_WORKERS]
        rs.sort(key=lambda t: -t[1]["annual"])
        return rs

    out_dirs = [REPO / "nenshu" / "shokushu", REPO / "nenshu" / "pref"]
    for d in out_dirs:
        d.mkdir(parents=True, exist_ok=True)
    urls = []

    # --- 職種ページ
    for n in keep:
        sx = occ[n]["男女計"]
        a = sx["all"]
        ann, cash, bonus = a["annual"], a["cash"], a["bonus"]
        sid = slug[n]
        path = "/nenshu/shokushu/%s/" % sid
        r = rank[n]
        diff = int(round((ann - nat_annual) / 10.0))
        sents = []
        sents.append("令和7年賃金構造基本統計調査によると、%sの平均年収は約%s万円です。内訳は、毎月のきまって支給する現金給与額が約%s万円、"
                     "年間賞与その他特別給与額が約%s万円で、年収は「月額×12+賞与」で計算しています。"
                     % (n, man(ann), man1(cash), man(bonus)))
        if abs(diff) >= 1:
            sents.append("全職種の平均(約%s万円)と比べると%s万円%s水準で、掲載している%d職種のうち%d位です。"
                         % (man(nat_annual), "{:,}".format(abs(diff)), "高い" if diff > 0 else "低い", N, r))
        else:
            sents.append("全職種の平均(約%s万円)とほぼ同じ水準で、掲載している%d職種のうち%d位です。" % (man(nat_annual), N, r))
        m = occ[n].get("男", {}).get("all")
        f = occ[n].get("女", {}).get("all")
        if (m and f and m["annual"] and f["annual"] and (m["wk"] or 0) >= MIN_SEX_WORKERS
                and (f["wk"] or 0) >= MIN_SEX_WORKERS):
            sd = int(round((m["annual"] - f["annual"]) / 10.0))
            sents.append("男女別では男性が約%s万円、女性が約%s万円で、差は約%s万円%s。"
                         % (man(m["annual"]), man(f["annual"]), "{:,}".format(abs(sd)),
                            "(男性が高い)" if sd > 0 else ("(女性が高い)" if sd < 0 else "")))
        alist = [(lab, d) for lab, d in ages5.get(n, []) if d["wk"] and d["wk"] >= MIN_AGE_WORKERS]
        if len(alist) >= 3:
            peak = max(alist, key=lambda x: x[1]["annual"])
            youngest = alist[0] if not alist[0][0].startswith(("～", "~")) else (alist[1] if len(alist) > 1 else alist[0])
            if youngest[0] != peak[0]:
                sents.append("年齢階級別では%sが約%s万円で最も高く、最も若い階級の%sは約%s万円です。"
                             % (age_label(peak[0]), man(peak[1]["annual"]), age_label(youngest[0]), man(youngest[1]["annual"])))
            else:
                sents.append("年齢階級別では%sが約%s万円で最も高くなっています。" % (age_label(peak[0]), man(peak[1]["annual"])))
        sents.append("平均年齢は%s歳、平均勤続年数は%s年、月の所定内実労働時間は約%d時間・超過実労働時間は約%d時間です。"
                     % (a["age"], a["ten"], a["hrs"], a["ot"]))
        sz = {k: sx[k] for k in ("l1000", "m100", "s10")}
        big, small = sz["l1000"], sz["s10"]
        if (big["annual"] and small["annual"] and (big["wk"] or 0) >= MIN_SIZE_WORKERS and (small["wk"] or 0) >= MIN_SIZE_WORKERS):
            bd = int(round((big["annual"] - small["annual"]) / 10.0))
            sents.append("企業規模別では、従業員1,000人以上で約%s万円、10〜99人で約%s万円と、約%s万円の開きがあります。"
                         % (man(big["annual"]), man(small["annual"]), "{:,}".format(abs(bd))))

        # 表
        t_main = table(["区分", "平均年収(万円)", "月額(万円)", "賞与等(万円)", "労働者数(人)"], [
            ["男女計", man(ann), man1(cash), man(bonus), "{:,}".format(int(a["wk"]))]] + [
            [lab, man(d["annual"]), man1(d["cash"]), man(d["bonus"]), "{:,}".format(int(d["wk"]))]
            for lab, d in (("男性", m), ("女性", f)) if d and d["annual"] and (d["wk"] or 0) >= MIN_SEX_WORKERS],
            num_cols=(1, 2, 3, 4))
        t_size = table(["企業規模", "平均年収(万円)", "労働者数(人)"], [
            [lab, man(d["annual"]), "{:,}".format(int(d["wk"]))]
            for lab, d in (("1,000人以上", sz["l1000"]), ("100〜999人", sz["m100"]), ("10〜99人", sz["s10"]))
            if d["annual"] and (d["wk"] or 0) >= MIN_SIZE_WORKERS], num_cols=(1, 2))
        t_age = table(["年齢階級", "平均年収(万円)", "労働者数(人)"], [
            [age_label(lab), man(d["annual"]), "{:,}".format(int(d["wk"]))] for lab, d in alist], num_cols=(1, 2)) \
            if len(alist) >= 3 else ""

        near = [x for x in keep[max(0, r - 4):r + 3] if x != n][:6]
        rel = "".join('<li><a href="/nenshu/shokushu/%s/">%s(約%s万円)</a></li>'
                      % (slug[x], esc(x), man(occ[x]["男女計"]["all"]["annual"])) for x in near)

        aff, matched = pick_affiliates(n, jobs)

        def occ_pref_html(n):
            rs = po_rows_for(n)
            if len(rs) < 5:
                return ""
            hi, lo = rs[0], rs[-1]
            lead = ('<p>都道府県別(令和7年 都道府県別第3表)では、%sが約%s万円で最も高く、%sが約%s万円で最も低くなっています'
                    '(労働者数%s人以上の%d都道府県)。</p>'
                    % (PREFS[hi[0] - 1], man(hi[1]["annual"]), PREFS[lo[0] - 1], man(lo[1]["annual"]),
                       "{:,}".format(MIN_PO_TABLE_WORKERS), len(rs)))
            t = table(["順位", "都道府県", "平均年収(万円)", "月額(万円)", "労働者数(人)"], [
                [str(i + 1), po_link(c, n, PREFS[c - 1]), man(d["annual"]), man1(d["cash"]), "{:,}".format(int(d["wk"]))]
                for i, (c, d) in enumerate(rs)], num_cols=(0, 2, 3, 4))
            return '<h2 class="ct-h2">都道府県別の%sの年収</h2><div class="ct-body">%s</div>%s' % (esc(n), lead, t)

        body = ('<section class="ct-section" style="padding-top:var(--ct-s-4)"><p class="ct-eyebrow">令和7年 賃金構造基本統計調査</p>'
                '<h1 class="ct-h1">%sの平均年収は約%s万円【%d年版】</h1>'
                '<p class="ct-lead">月給約%s万円+賞与等約%s万円。男女別・年齢別・企業規模別の年収を公的統計で確認できます。</p>'
                '<div class="ct-body">%s</div>'
                '<h2 class="ct-h2">%sの年収データ</h2>%s'
                '%s'
                '<h2 class="ct-h2">企業規模別の年収</h2>%s'
                '%s'
                '%s%s'
                '%s'
                '<nav class="ct-related" aria-label="関連ページ"><h2 class="ct-h3">年収が近い職種</h2><ul>%s</ul>'
                '<ul style="margin-top:var(--ct-s-3)"><li><a href="/nenshu/">職種別・都道府県別の平均年収一覧</a></li></ul></nav>'
                '%s</section>'
                % (esc(n), man(ann), YEAR, man1(cash), man(bonus),
                   "".join("<p>%s</p>" % esc(s) for s in sents),
                   esc(n), t_main,
                   ("<h2 class=\"ct-h2\">年齢階級別の年収</h2>" + t_age) if t_age else "",
                   t_size, occ_pref_html(n), cta_diagnosis(sid), aff_html(sid, aff, matched),
                   ('<p class="ct-note" style="border:0;padding:0">本ページは転職サービスのアフィリエイト広告(PR)を含みます。</p>' if aff else ""),
                   rel, source_note()))
        title = "%sの平均年収は約%s万円｜%d年版 男女別・年齢別・企業規模別" % (n, man(ann), YEAR)
        desc = ("%sの平均年収は約%s万円(月給約%s万円+賞与等約%s万円)。平均年齢%s歳・勤続%s年。令和7年賃金構造基本統計調査から男女別・年齢別・企業規模別の年収を掲載。"
                % (n, man(ann), man1(cash), man(bonus), a["age"], a["ten"]))
        crumbs = [("年収UPシミュレーター", "/"), ("平均年収一覧", "/nenshu/"), (n, path)]
        wr(REPO / "nenshu" / "shokushu" / sid / "index.html", page_shell(path, title, desc, crumbs, body))
        urls.append((path, "0.6"))

    # --- 都道府県ページ
    pk = []
    for pf in PREFS:
        sh = short_pref(pf)
        d = pref_ref.get(sh, {}).get("男女計")
        if d and d["annual"]:
            pk.append(pf)
    pk.sort(key=lambda p: -pref_ref[short_pref(p)]["男女計"]["annual"])
    prank = {p: i + 1 for i, p in enumerate(pk)}
    P = len(pk)
    for pf in pk:
        sh = short_pref(pf)
        code = "%02d" % (PREFS.index(pf) + 1)
        path = "/nenshu/pref/%s/" % code
        rr = pref_ref[sh]
        d = rr["男女計"]
        ann, cash, bonus = d["annual"], d["cash"], d["bonus"]
        diff = int(round((ann - nat_annual) / 10.0))
        pr = prank[pf]
        sents = ["令和7年賃金構造基本統計調査によると、%sの平均年収は約%s万円です。内訳は、毎月のきまって支給する現金給与額が約%s万円、"
                 "年間賞与その他特別給与額が約%s万円で、年収は「月額×12+賞与」で計算しています。" % (pf, man(ann), man1(cash), man(bonus))]
        sents.append("全国平均(約%s万円)と比べると%s万円%s水準で、47都道府県のうち%d位です。"
                     % (man(nat_annual), "{:,}".format(abs(diff)), "高い" if diff > 0 else "低い", pr) if abs(diff) >= 1
                     else "全国平均(約%s万円)とほぼ同じ水準で、47都道府県のうち%d位です。" % (man(nat_annual), pr))
        mm, ff = rr.get("男"), rr.get("女")
        if mm and ff and mm["annual"] and ff["annual"]:
            sents.append("男女別では男性が約%s万円、女性が約%s万円で、差は約%s万円です。"
                         % (man(mm["annual"]), man(ff["annual"]), "{:,}".format(int(round((mm["annual"] - ff["annual"]) / 10.0)))))
        sents.append("平均年齢は%s歳、平均勤続年数は%s年、月の所定内実労働時間は約%d時間・超過実労働時間は約%d時間です。"
                     % (d["age"], d["ten"], d["hrs"], d["ot"]))
        al = [(lab, x) for lab, x in pref_ages.get(sh, []) if x["wk"] and x["wk"] >= MIN_PREF_AGE_WORKERS]
        if len(al) >= 4:
            peak = max(al, key=lambda x: x[1]["annual"])
            sents.append("年齢階級別では%sが約%s万円で最も高く、%sは約%s万円です。"
                         % (age_label(peak[0]), man(peak[1]["annual"]), age_label(al[0][0]), man(al[0][1]["annual"])))
        il = [(n_, x) for n_, x in pref_inds.get(sh, []) if x["annual"] and x["wk"] and x["wk"] >= MIN_PREF_IND_WORKERS]
        il.sort(key=lambda t: -t[1]["annual"])
        if len(il) >= 4:
            sents.append("産業別(大分類)では、%sが約%s万円で最も高く、%sが約%s万円で最も低くなっています。"
                         % (re.sub(r"^[A-Z]", "", il[0][0]), man(il[0][1]["annual"]),
                            re.sub(r"^[A-Z]", "", il[-1][0]), man(il[-1][1]["annual"])))
        t_main = table(["区分", "平均年収(万円)", "月額(万円)", "賞与等(万円)", "労働者数(人)"], [
            [lab, man(x["annual"]), man1(x["cash"]), man(x["bonus"]), "{:,}".format(int(x["wk"]))]
            for lab, x in (("男女計", d), ("男性", mm), ("女性", ff)) if x and x["annual"]], num_cols=(1, 2, 3, 4))
        t_age = table(["年齢階級", "平均年収(万円)", "労働者数(人)"], [
            [age_label(lab), man(x["annual"]), "{:,}".format(int(x["wk"]))] for lab, x in al], num_cols=(1, 2)) if len(al) >= 4 else ""
        t_ind = table(["産業(大分類)", "平均年収(万円)", "労働者数(人)"], [
            [esc(re.sub(r"^[A-Z]", "", n_)), man(x["annual"]), "{:,}".format(int(x["wk"]))] for n_, x in il],
            num_cols=(1, 2), first_wrap=True) if len(il) >= 4 else ""
        near = [x for x in pk[max(0, pr - 4):pr + 3] if x != pf][:6]
        rel = "".join('<li><a href="/nenshu/pref/%02d/">%s(約%s万円)</a></li>'
                      % (PREFS.index(x) + 1, esc(x), man(pref_ref[short_pref(x)]["男女計"]["annual"])) for x in near)
        top_occ = keep[:6]
        occ_rel = "".join('<li><a href="/nenshu/shokushu/%s/">%s(全国平均 約%s万円)</a></li>'
                          % (slug[x], esc(x), man(occ[x]["男女計"]["all"]["annual"])) for x in top_occ)
        aff = top_affiliates(jobs)
        cnum = PREFS.index(pf) + 1
        ol = [(n_, po[cnum][n_]["男女計"]) for n_ in keep
              if (po.get(cnum, {}).get(n_, {}).get("男女計") or {}).get("wk", 0) >= MIN_PO_TABLE_WORKERS]
        ol.sort(key=lambda t: -t[1]["annual"])
        t_po = ('<h2 class="ct-h2" id="shokushu">%sの職種別の年収</h2>' % pf + table(
            ["職種", "平均年収(万円)", "全国との差(万円)", "労働者数(人)"],
            [[po_link(cnum, n_, n_), man(x["annual"]),
              "%+d" % int(round((x["annual"] - occ[n_]["男女計"]["all"]["annual"]) / 10.0)), "{:,}".format(int(x["wk"]))]
             for n_, x in ol], num_cols=(1, 2, 3), first_wrap=True)) if len(ol) >= 5 else ""
        body = ('<section class="ct-section" style="padding-top:var(--ct-s-4)"><p class="ct-eyebrow">令和7年 賃金構造基本統計調査</p>'
                '<h1 class="ct-h1">%sの平均年収は約%s万円【%d年版】</h1>'
                '<p class="ct-lead">全国%d位。男女別・年齢階級別・産業別の年収を公的統計で確認できます。</p>'
                '<div class="ct-body">%s</div>'
                '<h2 class="ct-h2">%sの年収データ</h2>%s'
                '%s%s%s'
                '%s%s'
                '<p class="ct-note" style="border:0;padding:0">本ページは転職サービスのアフィリエイト広告(PR)を含みます。</p>'
                '<nav class="ct-related" aria-label="関連ページ"><h2 class="ct-h3">年収が近い都道府県</h2><ul>%s</ul>'
                '<h2 class="ct-h3" style="margin-top:var(--ct-s-5)">全国の職種別 平均年収(上位)</h2><ul>%s</ul>'
                '<ul style="margin-top:var(--ct-s-3)"><li><a href="/nenshu/">職種別・都道府県別の平均年収一覧</a></li></ul></nav>'
                '%s</section>'
                % (pf, man(ann), YEAR, pr, "".join("<p>%s</p>" % esc(s) for s in sents), pf, t_main,
                   ("<h2 class=\"ct-h2\">年齢階級別の年収</h2>" + t_age) if t_age else "",
                   ("<h2 class=\"ct-h2\">産業別の年収</h2>" + t_ind) if t_ind else "", t_po,
                   cta_diagnosis("pref" + code), aff_html("pref" + code, aff, False), rel, occ_rel, source_note()))
        title = "%sの平均年収は約%s万円｜%d年版 男女別・年齢別・産業別" % (pf, man(ann), YEAR)
        desc = ("%sの平均年収は約%s万円(全国%d位、全国平均より%s万円%s)。令和7年賃金構造基本統計調査から男女別・年齢階級別・産業別の年収を掲載。"
                % (pf, man(ann), pr, "{:,}".format(abs(diff)), "高い" if diff > 0 else "低い"))
        crumbs = [("年収UPシミュレーター", "/"), ("平均年収一覧", "/nenshu/"), (pf, path)]
        wr(REPO / "nenshu" / "pref" / code / "index.html", page_shell(path, title, desc, crumbs, body))
        urls.append((path, "0.6"))

    # --- 職種×都道府県ページ(段階公開分だけ)
    po_paths = []
    for (c, n) in po_cands:
        if (c, n) not in po_live:
            continue
        pf = PREFS[c - 1]
        x = po[c][n]["男女計"]
        nat_o = occ[n]["男女計"]["all"]
        path = po_path(c, n)
        rs = po_rows_for(n)
        prk = next(i + 1 for i, (cc, _) in enumerate(rs) if cc == c)
        dn = int(round((x["annual"] - nat_o["annual"]) / 10.0))
        pref_all = pref_ref.get(short_pref(pf), {}).get("男女計")
        sents = ["令和7年賃金構造基本統計調査(都道府県別第3表)によると、%sの%sの平均年収は約%s万円です。内訳は、毎月のきまって支給する現金給与額が約%s万円、"
                 "年間賞与その他特別給与額が約%s万円で、年収は「月額×12+賞与」で計算しています。"
                 % (pf, n, man(x["annual"]), man1(x["cash"]), man(x["bonus"]))]
        if abs(dn) >= 1:
            sents.append("全国の%s(約%s万円)と比べると%s万円%s、データのある%d都道府県のうち%d位です。"
                         % (n, man(nat_o["annual"]), "{:,}".format(abs(dn)), "高く" if dn > 0 else "低く", len(rs), prk))
        else:
            sents.append("全国の%s(約%s万円)とほぼ同じ水準で、データのある%d都道府県のうち%d位です。" % (n, man(nat_o["annual"]), len(rs), prk))
        if pref_all and pref_all["annual"]:
            dp = int(round((x["annual"] - pref_all["annual"]) / 10.0))
            sents.append("%sの全職種平均(約%s万円)と比べると%s。" % (
                pf, man(pref_all["annual"]), ("約%s万円%s" % ("{:,}".format(abs(dp)), "高い水準です" if dp > 0 else "低い水準です")) if abs(dp) >= 1 else "ほぼ同じ水準です"))
        m, f = po[c][n].get("男"), po[c][n].get("女")
        sx_ok = m and f and (m["wk"] or 0) >= MIN_PO_SEX_WORKERS and (f["wk"] or 0) >= MIN_PO_SEX_WORKERS
        if sx_ok:
            sents.append("男女別では男性が約%s万円、女性が約%s万円です。" % (man(m["annual"]), man(f["annual"])))
        sents.append("平均年齢は%s歳、平均勤続年数は%s年、月の所定内実労働時間は約%d時間・超過実労働時間は約%d時間で、推計の労働者数は約%s人です。"
                     % (x["age"], x["ten"], x["hrs"], x["ot"], "{:,}".format(int(x["wk"]))))
        t_main = table(["区分", "平均年収(万円)", "月額(万円)", "賞与等(万円)", "労働者数(人)"], [
            [lab, man(d["annual"]), man1(d["cash"]), man(d["bonus"]), "{:,}".format(int(d["wk"]))]
            for lab, d in (("男女計", x),) + ((("男性", m), ("女性", f)) if sx_ok else ())] + [
            ["(参考)全国の%s" % esc(n), man(nat_o["annual"]), man1(nat_o["cash"]), man(nat_o["bonus"]), "{:,}".format(int(nat_o["wk"]))]],
            num_cols=(1, 2, 3, 4))
        near = rs[max(0, prk - 4):prk + 3]
        t_near = table(["順位", "都道府県", "平均年収(万円)"], [
            [str(rs.index(t) + 1), ("<b>%s</b>" % esc(pf)) if t[0] == c else po_link(t[0], n, PREFS[t[0] - 1]), man(t[1]["annual"])]
            for t in near], num_cols=(0, 2))
        others = sorted([n_ for n_ in keep if (c, n_) in po_live and n_ != n],
                        key=lambda n_: abs(po[c][n_]["男女計"]["annual"] - x["annual"]))[:8]
        o_rel = "".join('<li><a href="%s">%s(約%s万円)</a></li>' % (po_path(c, n_), esc(n_), man(po[c][n_]["男女計"]["annual"])) for n_ in others)
        pid = "po%02d%s" % (c, slug[n])
        aff, matched = pick_affiliates(n, jobs)
        q1 = "%sの%sの平均年収はいくらですか?" % (pf, n)
        a1 = "令和7年賃金構造基本統計調査では約%s万円(月給約%s万円+賞与等約%s万円)です。" % (man(x["annual"]), man1(x["cash"]), man(x["bonus"]))
        q2 = "%sの%sの年収は全国と比べて高いですか?" % (pf, n)
        a2 = ("全国平均(約%s万円)より約%s万円%s、%d都道府県中%d位です。" % (man(nat_o["annual"]), "{:,}".format(abs(dn)), "高く" if dn > 0 else "低く", len(rs), prk)
              if abs(dn) >= 1 else "全国平均(約%s万円)とほぼ同じで、%d都道府県中%d位です。" % (man(nat_o["annual"]), len(rs), prk))
        faq = {"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": [
            {"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in ((q1, a1), (q2, a2))]}
        body = ('<section class="ct-section" style="padding-top:var(--ct-s-4)"><p class="ct-eyebrow">令和7年 賃金構造基本統計調査</p>'
                '<h1 class="ct-h1">%sの%sの平均年収は約%s万円【%d年版】</h1>'
                '<p class="ct-lead">全国の%sと比べて%s。%d都道府県中%d位。</p>'
                '<div class="ct-body">%s</div>'
                '<h2 class="ct-h2">%sの%sの年収データ</h2>%s'
                '<h2 class="ct-h2">近い順位の都道府県(%s)</h2>%s'
                '<p><a href="/nenshu/shokushu/%s/">%sの全国の年収(年齢別・企業規模別・47都道府県の比較)</a></p>'
                '<h2 class="ct-h2">よくある質問</h2><div class="ct-body"><h3 class="ct-h3">%s</h3><p>%s</p><h3 class="ct-h3">%s</h3><p>%s</p></div>'
                '%s%s%s'
                '<nav class="ct-related" aria-label="関連ページ"><h2 class="ct-h3">%sのほかの職種</h2><ul>%s</ul>'
                '<ul style="margin-top:var(--ct-s-3)"><li><a href="/nenshu/pref/%02d/">%sの平均年収(全職種)</a></li>'
                '<li><a href="/nenshu/">職種別・都道府県別の平均年収一覧</a></li></ul></nav>'
                '%s</section>'
                % (pf, esc(n), man(x["annual"]), YEAR,
                   esc(n), ("約%s万円%s" % ("{:,}".format(abs(dn)), "高い" if dn > 0 else "低い")) if abs(dn) >= 1 else "ほぼ同じ", len(rs), prk,
                   "".join("<p>%s</p>" % esc(t) for t in sents), pf, esc(n), t_main, esc(n), t_near,
                   slug[n], esc(n), esc(q1), esc(a1), esc(q2), esc(a2),
                   cta_diagnosis(pid), aff_html(pid, aff, matched),
                   ('<p class="ct-note" style="border:0;padding:0">本ページは転職サービスのアフィリエイト広告(PR)を含みます。</p>' if aff else ""),
                   pf, o_rel, c, pf, source_note()))
        title = "%sの%sの平均年収は約%s万円｜%d年版 全国比較・順位" % (pf, n, man(x["annual"]), YEAR)
        desc = ("%sの%sの平均年収は約%s万円(月給約%s万円+賞与等約%s万円)。全国の%sより%s、%d都道府県中%d位。令和7年賃金構造基本統計調査より。"
                % (pf, n, man(x["annual"]), man1(x["cash"]), man(x["bonus"]), n,
                   ("%s万円%s" % ("{:,}".format(abs(dn)), "高い" if dn > 0 else "低い")) if abs(dn) >= 1 else "ほぼ同じ", len(rs), prk))
        crumbs = [("年収UPシミュレーター", "/"), ("平均年収一覧", "/nenshu/"), (pf, "/nenshu/pref/%02d/" % c), (n, path)]
        wr(REPO / path.strip("/") / "index.html", page_shell(path, title, desc, crumbs, body, faq))
        urls.append((path, "0.5"))
        po_paths.append(path)
    # 公開から外れた組み合わせはゴミ箱へ(rollout を縮めたとき)
    keepset = {p_.strip("/") for p_ in po_paths}
    trash = Path.home() / ".Trash" / ("nenshu_po_" + TODAY)
    for d in sorted((REPO / "nenshu" / "pref").glob("[0-9][0-9]/j[0-9][0-9][0-9]")):
        rel_ = str(d.relative_to(REPO))
        if rel_ not in keepset:
            trash.mkdir(parents=True, exist_ok=True)
            os.rename(d, trash / rel_.replace("/", "_"))
    PO_BUILT.parent.mkdir(parents=True, exist_ok=True)
    PO_BUILT.write_text(json.dumps({"built": TODAY, "rollout": rollout, "candidates": len(po_cands), "pages": len(po_paths),
                                    "paths": po_paths}, ensure_ascii=False), encoding="utf-8")

    # --- 一覧ページ
    p_rows = [['<a href="/nenshu/pref/%02d/">%s</a>' % (PREFS.index(p) + 1, esc(p)), man(pref_ref[short_pref(p)]["男女計"]["annual"])]
              for p in pk]
    p_rows = [[str(i + 1)] + r for i, r in enumerate(p_rows)]
    o_rows = [[str(i + 1), '<a href="/nenshu/shokushu/%s/">%s</a>' % (slug[n], esc(n)), man(occ[n]["男女計"]["all"]["annual"])]
              for i, n in enumerate(keep)]
    path = "/nenshu/"
    top = keep[0]
    body = ('<section class="ct-section" style="padding-top:var(--ct-s-4)"><p class="ct-eyebrow">令和7年 賃金構造基本統計調査</p>'
            '<h1 class="ct-h1">職種別・都道府県別の平均年収一覧【%d年版】</h1>'
            '<p class="ct-lead">厚生労働省の賃金構造基本統計調査(令和7年)をもとに、%d職種と47都道府県の平均年収をまとめました。</p>'
            '<div class="ct-body"><p>%s</p></div>'
            '<h2 class="ct-h2" id="pref">都道府県別の平均年収ランキング</h2>%s'
            '<h2 class="ct-h2">職種別の平均年収ランキング</h2>%s'
            '%s'
            '<p class="ct-note" style="border:0;padding:0">転職サービスのリンクは、ページ内の職種別ページに掲載しています(アフィリエイト広告を含みます)。</p>'
            '%s</section>'
            % (YEAR, N,
               esc("全国の平均年収は約%s万円です。職種別では%sが約%s万円で最も高く、都道府県別では%sが約%s万円で最も高くなっています。年収は「きまって支給する現金給与額(月額)×12+年間賞与その他特別給与額」で算出しています。"
                   % (man(nat_annual), top, man(occ[top]["男女計"]["all"]["annual"]), pk[0], man(pref_ref[short_pref(pk[0])]["男女計"]["annual"]))),
               table(["順位", "都道府県", "平均年収(万円)"], p_rows, num_cols=(0, 2)),
               table(["順位", "職種", "平均年収(万円)"], o_rows, num_cols=(0, 2), first_wrap=True),
               cta_diagnosis("nenshu-index"), source_note()))
    title = "職種別・都道府県別の平均年収ランキング｜%d年版(令和7年賃金構造基本統計調査)" % YEAR
    desc = "令和7年賃金構造基本統計調査をもとに、%d職種と47都道府県の平均年収をランキングで掲載。気になる職種・地域の男女別・年齢別の年収も確認できます。" % N
    wr(REPO / "nenshu" / "index.html", page_shell(path, title, desc, [("年収UPシミュレーター", "/"), ("平均年収一覧", None)], body))
    urls.insert(0, (path, "0.8"))

    update_sitemap(urls)
    if do_link:
        add_internal_links()
    stats = dict(occupations_total=len(order), occupations_pages=N, pref_pages=P, pref_occ_pages=len(po_paths),
                 pref_occ_candidates=len(po_cands), index=1, total=len(urls))
    (REPO / "nenshu" / "_build_stats.json").write_text(json.dumps(stats, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False))


def wr(p, s):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(s, encoding="utf-8")


# ---------------------------------------------------------------- sitemap / 内部リンク

def update_sitemap(urls):
    p = REPO / "sitemap.xml"
    s = p.read_text(encoding="utf-8")
    s = re.sub(r"\s*<url>\s*<loc>https://nenshu\.clawtechjp\.com/nenshu/[^<]*</loc>.*?</url>", "", s, flags=re.S)
    add = "".join(
        "\n  <url>\n    <loc>%s%s</loc>\n    <lastmod>%s</lastmod>\n    <changefreq>yearly</changefreq>\n    <priority>%s</priority>\n  </url>"
        % (BASE, u, TODAY, pr) for u, pr in urls)
    s = s.replace("</urlset>", add + "\n</urlset>")
    p.write_text(s, encoding="utf-8")


LINK_MARK = "<!-- stat-link -->"


# ブログ記事のファイル名の頭 → 対応する職種ページ(SEO: 記事⇔統計ページを相互に結ぶ)
BLOG_JOBS = [
    (("acct-", "keiri-"), ["j056", "j033"]),
    (("care-", "kaigo-"), ["j070", "j071", "j030"]),
    (("constr-", "construction-", "kensetsu-"), ["j007", "j008"]),
    (("it-engineer-", "freelance-engineer-"), ["j011", "j010"]),
    (("hoikushi-",), ["j029"]),
    (("youchien-",), ["j035", "j029"]),
    (("kangoshi-", "nurse-"), ["j020", "j021"]),
    (("pharma-", "pharmacist-", "yakuzaishi-"), ["j017"]),
    (("rehabili-",), ["j024"]),
]
TOP_JOBS = ["j020", "j070", "j029", "j017", "j024", "j011", "j010", "j056", "j007", "j008", "j033", "j030"]
A = 'style="color:inherit"'


def _job_name(jid):
    f = REPO / "nenshu" / "shokushu" / jid / "index.html"
    if not f.exists():
        return None
    m = re.search(r"<title>(.+?)の平均年収", f.read_text(encoding="utf-8"))
    return m.group(1) if m else None


def _blog_title(f):
    m = re.search(r"<title>(.+?)</title>", f.read_text(encoding="utf-8"), flags=re.S)
    return html.unescape(m.group(1).split("｜")[0].split("|")[0].strip()) if m else f.stem


def _strip_mark(t, closing):
    return re.sub(r"\n?" + re.escape(LINK_MARK) + r".*?" + closing, "", t, flags=re.S)


def add_internal_links():
    pref_links = " / ".join('<a href="/nenshu/pref/%02d/" %s>%s</a>' % (i + 1, A, esc(short_pref(n)))
                            for i, n in enumerate(PREFS))
    top_jobs = [(j, _job_name(j)) for j in TOP_JOBS]
    job_links = " / ".join('<a href="/nenshu/shokushu/%s/" %s>%s</a>' % (j, A, esc(n)) for j, n in top_jobs if n)
    # トップページ: 人気職種と47都道府県へ直接リンク(クロールの入口)
    idx = REPO / "index.html"
    s = _strip_mark(idx.read_text(encoding="utf-8"), "</nav>")
    block = (LINK_MARK + '<nav aria-label="公的統計の年収データ" style="max-width:720px;margin:24px auto;padding:0 16px;font-size:14px;line-height:2">'
             '<h2 style="font-size:16px;margin:0 0 4px">職種別・都道府県別の平均年収(令和7年賃金構造基本統計調査)</h2>'
             '<p style="margin:0 0 8px"><b>職種別:</b> ' + job_links + ' / <a href="/nenshu/" ' + A + '>全142職種</a></p>'
             '<p style="margin:0"><b>都道府県別:</b> ' + pref_links + '</p></nav>')
    if "</body>" in s:
        idx.write_text(s.replace("</body>", block + "</body>", 1), encoding="utf-8")
    # ブログ記事 → 対応する職種ページ + 一覧。逆向き(職種ページ → 記事)も集める
    job_posts = {}
    for f in sorted((REPO / "blog").glob("*.html")):
        if f.name == "index.html":
            continue
        t = _strip_mark(f.read_text(encoding="utf-8"), "</p>")
        jobs = next((js for pre, js in BLOG_JOBS if f.name.startswith(pre)), [])
        links = ['<a href="/nenshu/shokushu/%s/">%sの平均年収(公的統計)</a>' % (j, esc(n)) for j in jobs for n in [_job_name(j)] if n]
        if "todofuken" in f.name or "chiiki" in f.name:
            links.append('<a href="/nenshu/#pref">都道府県別の平均年収ランキング</a>')
        links.append('<a href="/nenshu/">職種別・都道府県別の平均年収一覧</a>')
        for j in jobs:
            job_posts.setdefault(j, []).append((f.name, _blog_title(f)))
        blk = (LINK_MARK + '<p class="stat-link" style="max-width:720px;margin:24px auto;padding:0 16px;font-size:14px;line-height:2">'
               '令和7年賃金構造基本統計調査のデータ: ' + " / ".join(links) + '</p>')
        if "</article>" in t:
            t = t.replace("</article>", "</article>\n" + blk, 1)
        elif "</body>" in t:
            t = t.replace("</body>", blk + "\n</body>", 1)
        else:
            continue
        f.write_text(t, encoding="utf-8")
    # 職種ページ → 関連するブログ記事(最大6本)
    for d in sorted((REPO / "nenshu" / "shokushu").iterdir()):
        f = d / "index.html"
        if not f.exists():
            continue
        t = _strip_mark(f.read_text(encoding="utf-8"), "</section>")
        posts = job_posts.get(d.name, [])[:6]
        if posts:
            items = "".join('<li><a href="/blog/%s">%s</a></li>' % (n, esc(ti)) for n, ti in posts)
            t = t.replace("</main>", LINK_MARK + '<section><h2>この職種の年収・転職の解説記事</h2><ul>' + items + '</ul></section>\n</main>', 1)
        f.write_text(t, encoding="utf-8")

if __name__ == "__main__":
    build(sys.argv[1:])
