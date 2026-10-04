# -*- coding: utf-8 -*-
"""
航旅智慧营销决策台（国内航线为主、国际航线扩展）—— 最终定向升级版

闭环：市场数据 → 市场洞察 → 用户需求提取 → 动态旅客画像 → 航线/目的地匹配
      → 触达决策 → 营销渠道与时机 → AIGC营销内容 → 人工审核 → 效果评估 → 数据反馈 → 下一轮优化

四个页面保持不变：市场洞察 / 旅客画像 / 智能营销 / 效果评估
启动：python app.py   （依赖：pip install gradio pandas openpyxl plotly）
"""
import os, re, json, html, hashlib, datetime as dt
from html import escape as esc

import pandas as pd

DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
DATA_DIR = os.path.join(DIR, "data")
A_FILE = os.path.join(DATA_DIR, "A成员数据搜集表(1).md")          # 无Excel时的备用来源；优先读取 data/A成员数据搜集表*.xlsx
HOTWORD_FILE = os.path.join(DATA_DIR, "热词统计.md")                # 社交热词（辅助，可缺省）
BAIDU_XLSX = [os.path.join(DATA_DIR, "百度指数数据集.xlsx"), os.path.join(DIR, "数据集.xlsx")]  # 国际扩展验证数据（可缺省）
FEATURE_FILE = os.path.join(DATA_DIR, "目的地特征库.json")          # 产品/运营配置，非市场数据
COLORS = ["#f97316", "#2563eb", "#16a34a", "#9333ea", "#dc2626", "#0891b2", "#ca8a04", "#db2777", "#4b5563", "#65a30d"]
CARD = "border:1px solid #e5e7eb;border-radius:10px;padding:12px;flex:1;min-width:150px;background:#fff"
TH = 'style="border:1px solid #e5e7eb;padding:4px 8px;background:#f9fafb;text-align:left"'
TD = 'style="border:1px solid #e5e7eb;padding:4px 8px"'
SIM_NOTE = "Demo模拟数据，仅用于展示营销闭环效果，不代表东航真实业务数据。"
AIGC_NOTE = "AIGC生成内容，需人工审核后使用。"


# =====================================================================
# 1. 数据获取层（数据源适配器）——未来接入合法API/授权数据源时只替换这里
# =====================================================================
class LocalMarkdownSource:
    """读取项目 data/ 目录下的 A成员数据搜集表（Markdown表格）。"""
    name = "本地数据文件 data/A成员数据搜集表(1).md"

    def __init__(self, path=A_FILE):
        self.path = path

    def fetch(self):
        if not os.path.exists(self.path):
            raise FileNotFoundError(f"未找到数据文件：{self.path}")
        with open(self.path, encoding="utf-8") as f:
            return f.read()


def find_a_xlsx():
    """在 data/ 和程序目录下查找 A成员数据搜集表*.xlsx（优先使用Excel）。"""
    for d in (DATA_DIR, DIR):
        if os.path.isdir(d):
            fs = sorted(f for f in os.listdir(d) if f.startswith("A成员数据搜集表") and f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$"))
            if fs:
                return os.path.join(d, fs[0])
    return None


class LocalExcelSource:
    """直接读取本地 A成员数据搜集表.xlsx，每个工作表转为二维列表（键=工作表名，如 05_第二轮百度指数精筛）。"""
    def __init__(self, path):
        self.path = path
        self.name = "本地Excel " + os.path.basename(path)

    def fetch(self):
        book = pd.read_excel(self.path, sheet_name=None, header=None, dtype=object)
        out = {}
        for sn, df in book.items():
            rows = [["" if (v is None or (isinstance(v, float) and v != v)) else str(v).strip() for v in r] for r in df.values.tolist()]
            rows = [r for r in rows if any(r)]
            # 跳过表头前的标题行：取第一个非空单元格≥2个的行作为表头
            while rows and sum(1 for c in rows[0] if c) < 2:
                rows.pop(0)
            out[str(sn).strip()] = rows
        return out


# 例：未来可新增 class AuthorizedApiSource: def fetch(self): ...  （不做未授权爬虫）
DATA_SOURCE = LocalMarkdownSource()


def md_tables(text):
    """把 Markdown 文档按一级标题切分，每节解析 <table> 为二维列表。"""
    text = text.replace("\\", "")
    parts = re.split(r"^# (.+)$", text, flags=re.M)
    out = {}
    for i in range(1, len(parts), 2):
        rows = []
        for tr in re.findall(r"<tr>(.*?)</tr>", parts[i + 1], re.S):
            cells = [html.unescape(re.sub(r"<[^>]+>", "", c).strip())
                     for c in re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", tr, re.S)]
            rows.append(cells)
        out[parts[i].strip()] = rows
    return out


def current_source():
    """只使用最新版 A成员数据搜集表(1).md，不与旧版 Excel 混用。
    如需接入其他合法数据源（API/授权Excel），在此处返回新的 Source 对象即可。"""
    return DATA_SOURCE


def load_market_data(source=None):
    """数据获取：数据源 → 标准化（parse_destination_data）→ 第一轮粗筛 → 第二轮百度指数精筛 → 最终候选目的地池。"""
    src = source or current_source()
    STATE["source"] = src.name
    return src.fetch()


def _sec(tables, prefix):
    return next((v for k, v in tables.items() if k.startswith(prefix)), [])


def _df(rows):
    if not rows:
        return pd.DataFrame()
    head = [h for h in rows[0]]
    while head and not head[-1]:
        head.pop()
    body = [r + [""] * (len(head) - len(r)) for r in rows[1:] if len(r) >= 2 and r[0]]
    return pd.DataFrame([r[:len(head)] for r in body], columns=head)


def _num(x):
    try:
        v = float(str(x).strip())
        return v if v == v else None
    except Exception:
        return None


def parse_destination_data(text):
    """数据解析：把A成员数据表解析为第一轮、第二轮、最终候选、可行性证据、方法口径。"""
    T = text if isinstance(text, dict) else md_tables(text)
    r1 = _df(_sec(T, "03_"))
    r2 = _df(_sec(T, "05_"))
    final = _df(_sec(T, "06_"))
    final = final[final["目的地"].isin(r1["目的地"]) if not r1.empty else slice(None)] if not final.empty else final
    kw = _df(_sec(T, "04_"))
    ev = _df(_sec(T, "08_"))
    method = [r for r in _sec(T, "07_") if r and r[0]]
    raw = _df(_sec(T, "02_"))
    for c in ["名次分", "共识分", "可行性分", "第一轮粗筛总分", "直飞(0/1/2)", "签证(0/1/2)", "节假日适配(0/1)", "安全政策(0/1)"]:
        if c in r1:
            r1[c] = r1[c].map(_num)
    return dict(round1=r1, round2_raw=r2, final=final, keywords=kw, evidence=ev, method=method, raw=raw)


def trend_label(score, m):
    """趋势规则（07_方法与口径）：综合分≥55且动量>10%：快速升温；≥55：上升；45~55：平稳；<45：回落。"""
    if score is None:
        return "数据待补"
    if score >= 55 and (m or 0) > 0.10:
        return "快速升温"
    if score >= 55:
        return "上升"
    if score >= 45:
        return "平稳"
    return "回落"


def heat_from_values(recent3, overall, latest, prev):
    """与A数据表同口径：L=近3期均值/整体均值；M=最新/上期-1（限-1~1）；
    水平分=50L(0~100)；动量分=50+100M(0~100)；综合热度分=0.7×水平分+0.3×动量分。"""
    if not recent3 or not overall or not latest or not prev:
        return None
    L = recent3 / overall
    M = max(-1.0, min(1.0, latest / prev - 1))
    lv = max(0.0, min(100.0, 50 * L))
    mo = max(0.0, min(100.0, 50 + 100 * M))
    sc = 0.7 * lv + 0.3 * mo
    return dict(L=L, M=M, level=lv, momentum=mo, score=sc, trend=trend_label(sc, M))


def calculate_market_heat(parsed):
    """数据分析：按A表口径逐关键词复算热度，并聚合到目的地（0/缺失不作为有效值）。"""
    r2 = parsed["round2_raw"]
    rows = []
    for _, r in r2.iterrows():
        h = heat_from_values(_num(r.get("近3期均值")), _num(r.get("整体均值")), _num(r.get("最新期")), _num(r.get("上期期值")))
        rows.append(dict(目的地=r["目的地"], 关键词=r["关键词"], 近3期均值=_num(r.get("近3期均值")), 整体均值=_num(r.get("整体均值")),
                         最新期=_num(r.get("最新期")), 上期=_num(r.get("上期期值")),
                         水平比率L=None if not h else h["L"], 动量M=None if not h else h["M"],
                         水平分=None if not h else h["level"], 动量分=None if not h else h["momentum"],
                         综合热度分=None if not h else h["score"], 趋势=(h["trend"] if h else "数据待补"),
                         表内综合热度分=_num(r.get("综合热度分"))))
    kdf = pd.DataFrame(rows)
    agg = {}
    for d, g in kdf.groupby("目的地", sort=False):
        v = g.dropna(subset=["综合热度分"])
        if v.empty:
            agg[d] = dict(score=None, trend="数据待补", n=0, up=0)
            continue
        sc = v["综合热度分"].mean()
        mm = v["动量M"].mean()
        agg[d] = dict(score=sc, trend=trend_label(sc, mm), n=len(v), up=int(v["趋势"].isin(["快速升温", "上升"]).sum()), M=mm)
    return kdf, agg


def get_candidate_destinations(parsed, agg, top_n=10):
    """两轮筛选：第一轮按粗筛总分取TopN形成候选池（8~12）；第二轮仅在池内按百度指数综合热度排序。不设第三轮。"""
    r1 = parsed["round1"].copy()
    r1 = r1.dropna(subset=["第一轮粗筛总分"]).sort_values("第一轮粗筛总分", ascending=False, kind="stable")
    pool_names = list(parsed["final"]["目的地"]) if not parsed["final"].empty else []
    if not (8 <= len(pool_names) <= 12):  # 数据表未给最终池时，按口径自行取TopN
        pool_names = list(r1.head(top_n)["目的地"])
    fin = parsed["final"].set_index("目的地") if not parsed["final"].empty else pd.DataFrame()
    pool = []
    for d in pool_names:
        a = agg.get(d, {})
        r = r1[r1["目的地"] == d]
        r = r.iloc[0] if len(r) else {}
        f = fin.loc[d] if d in getattr(fin, "index", []) else {}
        sc = a.get("score")
        if sc is None:
            sc = _num(f.get("第二轮综合热度分")) if len(f) else None
        pool.append(dict(dest=d, r1=_num(r.get("第一轮粗筛总分")) if len(r) else None, heat=sc,
                         trend=a.get("trend") or (f.get("趋势") if len(f) else "数据待补"),
                         n=a.get("n", 0), up=a.get("up", 0),
                         air=_num(r.get("直飞(0/1/2)")) if len(r) else None, visa=_num(r.get("签证(0/1/2)")) if len(r) else None,
                         holiday=_num(r.get("节假日适配(0/1)")) if len(r) else None, safety=_num(r.get("安全政策(0/1)")) if len(r) else None,
                         table_heat=_num(f.get("第二轮综合热度分")) if len(f) else None,
                         basis=(f.get("选择依据") if len(f) else ""), src=(f.get("证据来源") if len(f) else "")))
    pool.sort(key=lambda x: -(x["heat"] or -1))
    for i, p in enumerate(pool):
        p["rank"] = i + 1
    return pool, r1


def load_features():
    try:
        with open(FEATURE_FILE, encoding="utf-8") as f:
            d = json.load(f)
        return {k: v for k, v in d.items() if not k.startswith("_")}
    except Exception:
        return {}


def load_hotwords():
    """热词统计：社交/平台热词，仅作辅助参考，不与百度指数混为同一指标。"""
    if not os.path.exists(HOTWORD_FILE):
        return {}
    T = md_tables(open(HOTWORD_FILE, encoding="utf-8").read())
    hits = {}
    for plat, rows in T.items():
        for r in rows:
            for c in r[1:]:
                if c and not c.isdigit() and len(c) <= 6:
                    hits.setdefault(c, set()).add(plat)
    return hits


def load_baidu_xlsx():
    """国际扩展验证数据：百度指数数据集.xlsx（每个sheet一个关键词，列：日期/指数）。0值视为无数据，不补值。"""
    p = next((x for x in BAIDU_XLSX if os.path.exists(x)), None)
    if not p:
        return {}
    out = {}
    for sn, s in pd.read_excel(p, sheet_name=None).items():
        s = s.iloc[:, :2].copy()
        s.columns = ["date", "v"]
        s["date"] = pd.to_datetime(s["date"].astype(str), format="%Y.%m.%d", errors="coerce")
        s = s.dropna(subset=["date"])
        s["v"] = pd.to_numeric(s["v"], errors="coerce")
        s["v"] = s["v"].where(s["v"] > 0)
        out[sn] = s.sort_values("date").reset_index(drop=True)
    return out


def xlsx_heat(series):
    v = series.dropna(subset=["v"])
    if len(v) < 4:
        return None
    h = heat_from_values(v["v"].tail(3).mean(), v["v"].mean(), v["v"].iloc[-1], v["v"].iloc[-2])
    if h:
        h.update(latest=int(v["v"].iloc[-1]), date=v["date"].iloc[-1])
    return h


# ---------------- 全局市场状态（由“更新市场数据”按钮刷新） ----------------
STATE = dict(ok=False, time="—", status="尚未加载", pool=[], round1=pd.DataFrame(), kdf=pd.DataFrame(), agg={},
             parsed={}, features={}, hot={}, xlsx={}, error="", source="", version=0)


def refresh_market_state():
    """读取 → 解析 → 计算热度 → 两轮候选池 → 写入全局状态。失败时保留旧状态并报告。"""
    try:
        text = load_market_data()
        parsed = parse_destination_data(text)
        kdf, agg = calculate_market_heat(parsed)
        pool, r1 = get_candidate_destinations(parsed, agg)
        if not pool:
            raise ValueError("候选目的地池为空")
        STATE.update(ok=True, parsed=parsed, kdf=kdf, agg=agg, pool=pool, round1=r1,
                     features=load_features(), hot=load_hotwords(), xlsx=load_baidu_xlsx(),
                     time=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), status="已加载最新数据", error="", version=STATE.get("version", 0) + 1)
    except Exception as e:
        STATE.update(ok=False, status="数据更新失败，请检查数据文件", error=str(e),
                     time=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    return STATE["ok"]


def known_places():
    names = set(STATE["features"])
    if not STATE["round1"].empty:
        names |= set(STATE["round1"]["目的地"])
    if not STATE["kdf"].empty:
        names |= set(STATE["kdf"]["目的地"])
    for sn in STATE["xlsx"]:
        for suf in ("旅游", "酒店", "签证", "机票"):
            if sn.endswith(suf) and len(sn) > len(suf) + 1:
                names.add(sn[: -len(suf)])
    names.discard("国际")
    return names


DOMESTIC_HINT = {"国内"}


def dest_type(d):
    f = STATE["features"].get(d, {})
    if f.get("type"):
        return "国内" if f["type"] == "国内" else "国际/地区"
    kw = STATE["parsed"].get("keywords") if STATE["parsed"] else None
    if kw is not None and not kw.empty and "目的地类型" in kw:
        r = kw[kw["目的地"] == d]
        if len(r):
            return "国内" if r.iloc[0]["目的地类型"] == "国内" else "国际/地区"
    return "未知"


# =====================================================================
# 2. 市场洞察页面
# =====================================================================
def status_html():
    col = "#16a34a" if STATE["ok"] else "#dc2626"
    err = f'<br><span style="font-size:12px;color:#dc2626">原因：{esc(STATE["error"])}</span>' if STATE["error"] else ""
    return (f'<div style="{CARD};max-width:none"><b>数据更新时间：</b>{STATE["time"]}　　'
            f'<b>数据状态：</b><span style="color:{col}">{STATE["status"]}</span>　　'
            f'<b>更新次数：</b>{STATE.get("version", 0)}　　<b>候选池：</b>{len(STATE["pool"])}个目的地　　<b>数据来源：</b>{esc(STATE.get("source") or DATA_SOURCE.name)}{err}</div>')


def fmt(v, n=2):
    return "—" if v is None or (isinstance(v, float) and v != v) else (f"{v:.{n}f}" if isinstance(v, float) else str(v))


def market_html():
    if not STATE["ok"] and not STATE["pool"]:
        return status_html() + "<p>暂无可用市场数据。请把数据文件放到 data/A成员数据搜集表(1).md 后点击【更新市场数据】。</p>"
    pool = STATE["pool"]
    dom = [p for p in pool if dest_type(p["dest"]) == "国内"]
    intl = [p for p in pool if dest_type(p["dest"]) != "国内"]
    tcol = {"快速升温": "#dc2626", "上升": "#f97316", "平稳": "#2563eb", "回落": "#6b7280"}
    cards = "".join(
        f'<div style="{CARD}"><div style="font-size:12px;color:#6b7280">第{p["rank"]}名 · {dest_type(p["dest"])}</div>'
        f'<div style="font-size:20px;font-weight:700">{esc(p["dest"])}</div>'
        f'<div>综合热度分 <b>{fmt(p["heat"])}</b></div>'
        f'<div style="color:{tcol.get(p["trend"], "#111")}">趋势：{p["trend"]}</div>'
        f'<div style="font-size:11px;color:#9ca3af">第一轮总分 {fmt(p["r1"], 0)}｜有效关键词 {p["n"]}</div></div>' for p in pool)
    top_dom = "、".join(p["dest"] for p in dom[:5]) or "无"
    rising = "、".join(p["dest"] for p in pool if p["trend"] in ("快速升温", "上升")) or "无"
    answer = (f'<div style="{CARD};max-width:none;background:#fff7ed"><b>现在市场上哪些目的地值得重点营销？</b><br>'
              f'国内主推（按第二轮综合热度）：<b>{top_dom}</b>；正在升温：<b>{rising}</b>；'
              f'国际/地区扩展候选：{"、".join(p["dest"] for p in intl) or "无"}（作为扩展场景，不作为默认推荐）。</div>')
    # 第一轮
    r1 = STATE["round1"]
    pool_set = {p["dest"] for p in pool}
    r1rows = "".join(
        f'<tr style="{"background:#ecfdf5" if r["目的地"] in pool_set else ""}"><td {TD}>{i+1}</td><td {TD}>{esc(r["目的地"])}</td>'
        f'<td {TD}>{fmt(r["名次分"],0)}</td><td {TD}>{fmt(r["共识分"],0)}</td><td {TD}>{fmt(r["直飞(0/1/2)"],0)}</td>'
        f'<td {TD}>{fmt(r["签证(0/1/2)"],0)}</td><td {TD}>{fmt(r["节假日适配(0/1)"],0)}</td><td {TD}>{fmt(r["安全政策(0/1)"],0)}</td>'
        f'<td {TD}><b>{fmt(r["第一轮粗筛总分"],0)}</b></td><td {TD}>{esc(str(r.get("可行性数据状态", "")))}</td>'
        f'<td {TD}>{"✔ 进入候选池" if r["目的地"] in pool_set else ""}</td></tr>'
        for i, (_, r) in enumerate(r1.head(20).iterrows()))
    r1tbl = (f'<table style="border-collapse:collapse;font-size:12px"><tr><th {TH}>序</th><th {TH}>目的地</th><th {TH}>名次分</th><th {TH}>共识分</th>'
             f'<th {TH}>航空</th><th {TH}>签证/入境</th><th {TH}>节假日</th><th {TH}>安全政策</th><th {TH}>粗筛总分</th><th {TH}>可行性数据状态</th><th {TH}>结果</th></tr>{r1rows}</table>'
             f'<p style="font-size:12px;color:#6b7280">共{len(r1)}个目的地参与第一轮，显示前20；绿色为进入第二轮的候选池（{len(pool)}个）。</p>')
    # 第二轮
    kdf = STATE["kdf"]
    kd = kdf[kdf["目的地"].isin(pool_set)]
    r2rows = ""
    for p in pool:
        for _, k in kd[kd["目的地"] == p["dest"]].iterrows():
            r2rows += (f'<tr><td {TD}>{esc(p["dest"])}</td><td {TD}>{esc(k["关键词"])}</td><td {TD}>{fmt(k["近3期均值"],1)}</td><td {TD}>{fmt(k["整体均值"],0)}</td>'
                       f'<td {TD}>{fmt(k["最新期"],0)}</td><td {TD}>{fmt(k["上期"],0)}</td><td {TD}>{fmt(k["水平比率L"])}</td><td {TD}>{fmt(k["动量M"])}</td>'
                       f'<td {TD}>{fmt(k["水平分"],1)}</td><td {TD}>{fmt(k["动量分"],1)}</td><td {TD}><b>{fmt(k["综合热度分"],1)}</b></td><td {TD}>{k["趋势"]}</td></tr>')
    r2tbl = (f'<details><summary>展开查看候选池内各关键词明细（{len(kd)}条）</summary><table style="border-collapse:collapse;font-size:12px">'
             f'<tr><th {TH}>目的地</th><th {TH}>关键词</th><th {TH}>近3期均值</th><th {TH}>整体均值</th><th {TH}>最新一期</th><th {TH}>上一期</th>'
             f'<th {TH}>L</th><th {TH}>M</th><th {TH}>水平分</th><th {TH}>动量分</th><th {TH}>综合热度分</th><th {TH}>趋势</th></tr>{r2rows}</table></details>')
    diffs = [abs(p["heat"] - p["table_heat"]) for p in pool if p["heat"] is not None and p["table_heat"] is not None]
    check = (f"系统按口径复算的目的地综合热度与A数据表“06_最终候选目的地”最大偏差 {max(diffs):.2f} 分（一致）。" if diffs and max(diffs) < 0.5
             else f"注意：复算结果与数据表存在偏差（最大 {max(diffs):.2f}），请核对原始数据。" if diffs else "")
    # 社交热词辅助
    hot = STATE["hot"]
    hotline = "；".join(f'{p["dest"]}：{"、".join(sorted(hot[p["dest"]]))}' for p in pool if p["dest"] in hot) or "暂无"
    # 方法与口径
    meth = "".join(f"<li>{esc(r[0])}{('：' + esc(r[1])) if len(r) > 1 and r[1] else ''}</li>" for r in STATE["parsed"].get("method", [])[:12])
    return (status_html() + answer +
            f'<h3>当前候选目的地（最终重点营销候选池，{len(pool)}个）</h3><div style="display:flex;gap:10px;flex-wrap:wrap">{cards}</div>'
            f'<h3>第一轮：粗筛结果（多平台榜单 + 平台共识 + 航空/签证/节假日/安全可行性）</h3>'
            f'<p style="font-size:13px">第一轮总分 = 名次分 + 共识分 + 航空可行性 + 签证/入境可行性 + 节假日适配 + 安全政策；目的是形成8~12个候选池，不判断最终谁最热门。</p>{r1tbl}'
            f'<h3>第二轮：百度指数精筛（仅对候选池）</h3>'
            f'<p style="font-size:13px">L=近3期均值÷整体均值；M=最新一期较上一期变化率（限-100%~+100%）；水平分=50×L；动量分=50+100×M；'
            f'<b>综合热度分=0.7×水平分+0.3×动量分</b>；趋势：综合分≥55且M>10%为快速升温，≥55上升，45~55平稳，&lt;45回落。目的地热度=有效关键词均值，0/缺失不计入。</p>'
            f'<p style="font-size:12px;color:#16a34a">{check}</p>{r2tbl}'
            f'<h3>社交平台热词（辅助参考，口径不同，不并入百度指数）</h3><p style="font-size:13px">{esc(hotline)}</p>'
            f'<details><summary>数据方法与口径（来自A数据表07节）</summary><ul style="font-size:12px">{meth}</ul></details>'
            f'<p style="color:#dc2626;font-size:12px">数据来源：A成员采集的公开榜单、百度指数公开数据与可行性证据（详见数据表01/08节，含链接与核验日期）；不代表东航真实客户数据。数据文件中标注“数据待补”的项不做编造。</p>')


def market_fig():
    pool = STATE["pool"]
    if not pool:
        return None
    try:
        import plotly.graph_objects as go
        xs = [p["dest"] for p in pool]
        fig = go.Figure()
        fig.add_bar(x=xs, y=[p["heat"] for p in pool], name="第二轮综合热度分", marker_color=[COLORS[0] if dest_type(d) == "国内" else COLORS[1] for d in xs],
                    text=[p["trend"] for p in pool], textposition="outside")
        fig.add_scatter(x=xs, y=[p["r1"] for p in pool], name="第一轮粗筛总分（右轴）", yaxis="y2", mode="lines+markers", line=dict(color=COLORS[2]))
        fig.update_layout(template="plotly_white", height=400, margin=dict(l=50, r=50, t=40, b=40), legend=dict(orientation="h", y=1.15),
                          yaxis=dict(title="综合热度分", rangemode="tozero"), yaxis2=dict(title="粗筛总分", overlaying="y", side="right", rangemode="tozero", showgrid=False))
        return fig
    except ImportError:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "PingFang SC", "Arial Unicode MS", "DejaVu Sans"]
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar([p["dest"] for p in pool], [p["heat"] or 0 for p in pool], color=COLORS[0])
        ax.set_ylabel("综合热度分")
        fig.tight_layout()
        return fig


TCOL = {"快速升温": "#dc2626", "上升": "#f97316", "平稳": "#2563eb", "回落": "#6b7280"}


def final_table():
    """06_最终候选目的地：按表内“最终顺序/第二轮综合热度分”排序，直接使用表内数值。"""
    f = STATE["parsed"].get("final") if STATE["parsed"] else None
    if f is None or f.empty:
        return pd.DataFrame(columns=["排名", "目的地", "综合热度分", "趋势"])
    sc = f["第二轮综合热度分"].map(_num)
    od = f["最终顺序"].map(_num) if "最终顺序" in f else None
    d = pd.DataFrame({"目的地": f["目的地"], "综合热度分": sc, "趋势": f.get("趋势", ""), "_o": od if od is not None else -sc})
    d = d.sort_values(["_o"] if od is not None else ["综合热度分"], ascending=od is not None, kind="stable").drop(columns="_o").reset_index(drop=True)
    d.insert(0, "排名", range(1, len(d) + 1))
    d["综合热度分"] = d["综合热度分"].map(lambda v: round(v, 2) if v is not None else None)
    return d


def top3_html():
    head = status_html() + f'<div style="font-size:12px;color:#6b7280;margin:4px 0">数据来源：{esc(STATE.get("source") or "—")}（06_最终候选目的地）</div>'
    d = final_table()
    if d.empty:
        return head + "<p>暂无最终候选目的地数据，请检查 data/ 下的 A成员数据搜集表。</p>"
    mx = max([v for v in d["综合热度分"] if v is not None] + [1])
    rows = ""
    for _, r in d.head(3).iterrows():
        c = TCOL.get(r["趋势"], "#6b7280")
        w = 0 if r["综合热度分"] is None else r["综合热度分"] / mx * 100
        rows += (f'<div style="margin:14px 0"><div style="display:flex;justify-content:space-between;align-items:center">'
                 f'<b style="font-size:16px">{r["排名"]}. {esc(r["目的地"])} <span style="font-size:12px;color:{c};border:1px solid {c};border-radius:4px;padding:0 4px">{esc(str(r["趋势"]))}</span></b>'
                 f'<span style="font-size:13px;color:#374151">最终热度分 <b>{fmt(r["综合热度分"])}</b></span></div>'
                 f'<div style="background:#f3f4f6;border-radius:6px;height:10px;margin-top:6px"><div style="width:{w:.1f}%;background:#f97316;height:10px;border-radius:6px"></div></div>'
                 f'<div style="font-size:12px;color:#6b7280;margin-top:3px">趋势：{esc(str(r["趋势"]))}</div></div>')
    return head + f'<div style="{CARD};max-width:none"><h3 style="margin:0">目的地热度排行（前三名）</h3>{rows}</div>'


def keyword_detail_html(dest):
    """05_第二轮百度指数精筛：列出所选目的地各关键词的综合热度分与趋势（表内数值）。"""
    r2 = STATE["parsed"].get("round2_raw") if STATE["parsed"] else None
    if not dest:
        return f'<div style="{CARD};max-width:none"><h3 style="margin:0">目的地关键词明细</h3><p style="color:#6b7280">请在上方“全部候选目的地”中点击一个目的地查看。</p></div>'
    if r2 is None or r2.empty:
        return "<p>暂无 05_第二轮百度指数精筛 数据。</p>"
    g = r2[r2["目的地"] == dest]
    lines = ""
    for _, r in g.iterrows():
        v = _num(r.get("综合热度分"))
        t = str(r.get("趋势", "")) or "数据待补"
        c = TCOL.get(t, "#6b7280")
        lines += (f'<div style="display:flex;justify-content:space-between;padding:8px 4px;border-bottom:1px solid #f3f4f6">'
                  f'<span>▮ {esc(str(r["关键词"]))}</span><span>综合热度分 <b>{fmt(v) if v is not None else "数据待补"}</b>　'
                  f'趋势 <span style="color:{c}">{esc(t)}</span></span></div>')
    if not lines:
        lines = "<p>该目的地在05表中没有关键词数据。</p>"
    return (f'<div style="{CARD};max-width:none"><h3 style="margin:0">目的地关键词明细 · {esc(dest)}</h3>'
            f'<div style="font-size:12px;color:#6b7280;margin:4px 0">来源：05_第二轮百度指数精筛</div>{lines}</div>')


def default_dest():
    d = final_table()
    return d.iloc[0]["目的地"] if len(d) else None


def refresh_market_ui():
    """【数据更新】：重新读取本地A成员数据搜集表，三部分全部刷新。"""
    refresh_market_state()
    d0 = default_dest()
    return top3_html(), final_table(), keyword_detail_html(d0), d0, status_html()


# =====================================================================
# 3. 用户需求提取（自然语言 → 结构化字段，含原文依据与置信度）
# =====================================================================
INTEREST_KW = {
    "美食": ["美食", "吃", "小吃", "餐厅", "火锅"], "购物": ["购物", "买买买", "免税", "逛街", "商场"],
    "娱乐": ["娱乐", "演出", "乐园", "夜生活", "游乐"], "自然风景": ["自然", "风景", "山水", "爬山", "看海", "徒步"],
    "历史人文": ["历史", "人文", "古城", "博物馆", "古迹"], "海岛": ["海岛", "海边", "沙滩"],
    "亲子": ["亲子", "孩子", "小孩", "带娃", "儿童"], "摄影": ["拍照", "摄影", "打卡"], "咖啡": ["咖啡"],
    "潮流": ["潮流", "时尚"], "商务": ["商务", "出差", "开会"], "短途旅行": ["短途", "周末"],
}
TAGS = ["美食", "购物", "娱乐", "自然风景", "历史人文", "海岛", "亲子", "家庭", "摄影", "咖啡", "潮流", "商务", "学生", "短途旅行"]
CN_NUM = {"一": 1, "两": 2, "二": 2, "俩": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
HOLIDAYS = {"国庆": (10, 1), "十一": (10, 1), "元旦": (1, 1), "五一": (5, 1), "劳动节": (5, 1), "春节": (2, 6), "暑假": (7, 15), "寒假": (1, 20), "中秋": (9, 25)}
DEPART_CITIES = ["上海", "北京", "广州", "深圳", "杭州", "南京", "成都", "重庆", "武汉", "西安", "昆明", "青岛", "厦门", "合肥", "济南", "宁波", "苏州", "天津", "长沙", "郑州"]
PREF_KW = {"交通方便": ["交通方便", "交通便利", "地铁近", "位置方便"], "减少步行": ["少走路", "不想走太多", "走路太多", "步行少", "少走"],
           "住宿品质": ["好一点的酒店", "高档酒店", "五星"], "性价比": ["性价比", "划算", "省钱"], "直飞": ["直飞"]}
EXCL_KW = {"不要太早的航班": ["太早的航班", "早班机", "太早的飞机", "不想早起"], "不要红眼航班": ["红眼"], "不要中转": ["不要中转", "不想中转", "不转机"],
           "不要跟团": ["不跟团", "不要跟团", "不想跟团"]}


def _ev(text, m):
    return f"“{m}”" if m else "—"


def holiday_date(name, today=None):
    today = today or dt.date.today()
    mth, day = HOLIDAYS[name]
    d = dt.date(today.year, mth, day)
    if name in ("国庆", "十一") and dt.date(today.year, 10, 1) <= today <= dt.date(today.year, 10, 7):
        return today
    return d if d >= today - dt.timedelta(days=3) else dt.date(today.year + 1, mth, day)


def analyze_user_need(text, form=None):
    """自然语言需求提取。form 为结构化表单（可缺省）；自然语言提取到的字段优先，表单作为补充。"""
    t = str(text or "").strip()
    form = form or {}
    F = {}

    def put(k, v, ev, conf):
        F[k] = dict(value=v, evidence=ev, conf=conf)

    places = known_places()
    found = []
    for p in sorted(places, key=len, reverse=True):
        for m in re.finditer(re.escape(p), t):
            if not any(m.start() < e and m.end() > s for s, e, _ in found):
                found.append((m.start(), m.end(), p))
    found.sort()
    # 出发城市
    dep = None
    m = re.search(r"从(.{2,4}?)(出发|飞|去|到|走)", t)
    if m:
        dep, dep_ev = m.group(1), m.group(0)
    else:
        for s, e, p in found:
            if re.match(r"(出发|去|飞往|飞|到)", t[e:e + 2]) and p in DEPART_CITIES:
                dep, dep_ev = p, t[s:e + 1]
                break
    if dep:
        put("出发城市", dep, _ev(t, dep_ev), "高")
    elif form.get("city"):
        put("出发城市", form["city"], "结构化表单", "中")
    else:
        put("出发城市", "未指定", "—", "低")
    # 目的地
    unknown = re.search(r"(不知道去哪|不确定去哪|没想好去哪|去哪(里)?好|推荐.*目的地|随便去)", t)
    cmp_ = re.search(r"(.{2,4})和(.{2,4})(不知道选|选哪|哪个好)", t)
    dests = [(s, e, p) for s, e, p in found if p != dep]
    # 国家名与城市同时出现时取城市
    names = [p for _, _, p in dests]
    names = [n for n in names if not any(n == STATE["features"].get(o, {}).get("country") for o in names if o != n)]
    if unknown and not names:
        put("目的地", "未指定", _ev(t, unknown.group(0)), "低")
    elif names:
        uniq = list(dict.fromkeys(names))
        put("目的地", "、".join(uniq), _ev(t, uniq[0]), "高" if len(uniq) == 1 else "中")
    elif form.get("dest") and form["dest"] not in ("未指定（系统推荐）", ""):
        put("目的地", form["dest"], "结构化表单", "中")
    else:
        put("目的地", "未指定", "原文未出现目的地", "低")
    F["目的地"]["compare"] = bool(cmp_) or len(names) >= 2
    # 人数
    m = re.search(r"(\d+|[一两二俩三四五六七八九十])\s*(个人|人|位|口)", t)
    m2 = re.search(r"一家(三|四|五)口", t)
    if m2:
        put("出行人数", CN_NUM[m2.group(1)], _ev(t, m2.group(0)), "高")
    elif m:
        n = int(m.group(1)) if m.group(1).isdigit() else CN_NUM[m.group(1)]
        put("出行人数", n, _ev(t, m.group(0)), "高")
    elif re.search(r"(我们俩|情侣|和男朋友|和女朋友|和老公|和老婆)", t):
        put("出行人数", 2, _ev(t, re.search(r"(我们俩|情侣|和男朋友|和女朋友|和老公|和老婆)", t).group(0)), "中")
    elif form.get("n"):
        put("出行人数", int(form["n"]), "结构化表单", "中")
    else:
        put("出行人数", "未指定", "—", "低")
    # 预算
    m = re.search(r"预算\D{0,4}?(\d+(?:\.\d+)?)\s*(万|千|k|K)?\s*(元)?\s*(/人|每人|一人|人均)?", t) or \
        re.search(r"(人均|每人)\s*(\d+(?:\.\d+)?)\s*(元)?", t)
    if m:
        g = m.groups()
        if m.re.pattern.startswith("(人均"):
            val, unit, per = float(g[1]), None, "/人"
        else:
            val, unit, per = float(g[0]), g[1], g[3]
        val *= 10000 if unit == "万" else 1000 if unit in ("千", "k", "K") else 1
        put("预算", int(val), _ev(t, m.group(0).strip()), "高" if per else "中")
        F["预算"]["per"] = "元/人" if per else "元（未注明是否人均，按人均处理）"
    elif form.get("budget"):
        put("预算", int(form["budget"]), "结构化表单", "中")
        F["预算"]["per"] = "元/人"
    else:
        put("预算", "未指定", "—", "低")
    # 日期
    m = re.search(r"(20\d\d)[-年./](\d{1,2})[-月./](\d{1,2})", t)
    m1 = re.search(r"(\d{1,2})月(\d{1,2})[日号]", t)
    m3 = re.search(r"(\d{1,2})月(份)?", t)
    hol = next((h for h in HOLIDAYS if h in t), None)
    hol_form = False
    if not hol and form.get("holiday") in HOLIDAYS:  # 营销任务中选择的假期
        hol, hol_form = form["holiday"], True
    today = dt.date.today()
    if m:
        d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        put("出行日期", d.isoformat(), _ev(t, m.group(0)), "高")
    elif m1:
        d = dt.date(today.year, int(m1.group(1)), int(m1.group(2)))
        d = d if d >= today else dt.date(today.year + 1, d.month, d.day)
        put("出行日期", d.isoformat(), _ev(t, m1.group(0)), "高")
    elif hol:
        put("出行日期", f"{hol}（约{holiday_date(hol).isoformat()}）", "营销任务中选择的假期" if hol_form else _ev(t, hol), "中")
        d = holiday_date(hol)
    elif m3 and 1 <= int(m3.group(1)) <= 12:
        mm = int(m3.group(1))
        d = dt.date(today.year if mm >= today.month else today.year + 1, mm, 15 if mm != today.month else max(today.day, 15))
        put("出行日期", f"{mm}月（按{d.isoformat()}估算）", _ev(t, m3.group(0)), "中")
    elif form.get("date"):
        try:
            d = dt.date.fromisoformat(str(form["date"]).strip())
            put("出行日期", d.isoformat(), "结构化表单", "中")
        except Exception:
            d = None
            put("出行日期", "未指定", "表单日期格式无法识别", "低")
    else:
        d = None
        put("出行日期", "未指定", "—", "低")
    F["出行日期"]["date"] = d if F["出行日期"]["value"] != "未指定" else None
    F["出行日期"]["holiday"] = hol
    # 兴趣
    ints, evs = [], []
    for k, kws in INTEREST_KW.items():
        h = next((w for w in kws if w in t), None)
        if h:
            ints.append(k)
            evs.append(h)
    for x in form.get("tags") or []:
        if x not in ints:
            ints.append(x)
    put("兴趣", ints or "未指定", "、".join(f"“{e}”" for e in evs) or ("结构化表单" if ints else "—"), "高" if evs else "中" if ints else "低")
    prefs = [(k, next(w for w in v if w in t)) for k, v in PREF_KW.items() if any(w in t for w in v)]
    put("偏好", [p for p, _ in prefs] or "未提及", "、".join(f"“{w}”" for _, w in prefs) or "—", "高" if prefs else "低")
    excl = [(k, next(w for w in v if w in t)) for k, v in EXCL_KW.items() if any(w in t for w in v)]
    put("排除条件", [p for p, _ in excl] or "未提及", "、".join(f"“{w}”" for _, w in excl) or "—", "高" if excl else "低")
    F["_text"] = t
    return F


# =====================================================================
# 4. 决策阶段 / 情绪与意向 / 触达决策（严格优先级）
# =====================================================================
REFUSE_KW = ["不要联系", "不要推荐", "拒绝营销", "不需要推荐", "不需要营销", "不要营销", "别推", "别发", "不要再发", "不要发", "屏蔽",
             "不想收到", "不希望收到", "别打扰", "别联系", "停止推送", "停止营销", "拒收", "不要给我发", "不要打电话", "退订", "不要任何营销"]
NOPLAN_PAT = [r"确定不去", r"不去了", r"不打算去", r"不想去了", r"取消(了)?(行程|出行|旅行)", r"没有.{0,6}计划", r"没计划", r"暂时不考虑", r"暂不考虑",
              r"不考虑出行", r"暂时不去", r"不出行了", r"暂无出行", r"没时间去"]
STRONG_KW = ["准备订", "准备买", "正在看机票", "正在看酒店", "在看机票", "订票", "下单", "买票", "确定去", "确定.*去", "已经定", "订好", "帮我订", "帮我推荐机票", "马上", "准备.*出行", "准备出发"]
HESITATE_KW = ["再看看", "再考虑", "还没决定", "没决定", "犹豫", "纠结", "观望", "再想想", "再等等", "还没想好", "不确定", "要不要", "比较一下"]
PRICE_KW = ["太贵", "贵", "价格高", "超预算", "预算不够", "预算有限", "便宜点", "涨价", "没多少钱", "穷游"]
VISA_KW = ["签证", "入境", "护照", "办证", "通行证"]
TIME_KW = ["请不了假", "时间不够", "没时间", "假期太短", "时间紧"]
CONV_KW = ["太远", "不方便", "转机", "中转", "太早", "折腾", "走太多"]
POS_KW = ["想去", "很想", "喜欢", "期待", "心动", "感兴趣", "开心", "向往", "必须去"]


def _hits(t, kws):
    return [k for k in kws if re.search(k, t)]


def analyze_emotion(text, need):
    t = re.sub(r"去不去", "", str(text or ""))  # “去不去还没决定”不是“不去”
    e = dict(refuse=_hits(t, REFUSE_KW), noplan=_hits(t, NOPLAN_PAT), strong=_hits(t, STRONG_KW), hesitate=_hits(t, HESITATE_KW),
             price=_hits(t, PRICE_KW), visa=_hits(t, VISA_KW), time=_hits(t, TIME_KW), conv=_hits(t, CONV_KW), pos=_hits(t, POS_KW))
    # “签证”只有带担忧语气才算顾虑
    if e["visa"] and not re.search(r"(签证|入境|护照|通行证).{0,6}(麻烦|难|担心|不清楚|不了解|来不及|怎么办|复杂)|(担心|不清楚|不了解).{0,6}(签证|入境)", t):
        e["visa"] = []
    # 预算明显不足 → 价格约束（来自结构化预算，而非情绪词）
    e["budget_tight"] = False
    concerns = []
    if e["price"]: concerns.append(("价格顾虑", "“" + e["price"][0] + "”"))
    if e["visa"]: concerns.append(("签证顾虑", "“" + e["visa"][0] + "”"))
    if e["time"]: concerns.append(("时间顾虑", "“" + e["time"][0] + "”"))
    if e["conv"]: concerns.append(("便利性顾虑", "“" + e["conv"][0] + "”"))
    if e["hesitate"]: concerns.append(("犹豫", "“" + e["hesitate"][0] + "”"))
    e["concerns"] = concerns
    if e["refuse"]:
        e["mood"] = "明确拒绝营销"
    elif e["noplan"]:
        e["mood"] = "暂无出行计划/明确取消"
    elif e["pos"] and concerns:
        e["mood"] = "谨慎（有意愿＋顾虑）"
    elif e["pos"] or e["strong"]:
        e["mood"] = "积极意向"
    elif concerns:
        e["mood"] = "犹豫/负面因素"
    else:
        e["mood"] = "中性"
    e["intent"] = "无" if (e["refuse"] and not e["pos"] and not e["strong"]) or e["noplan"] else "有" if (e["pos"] or e["strong"]) else "不明确"
    return e


def detect_stage(text, need, emo):
    t = str(text or "")
    dest = need["目的地"]["value"]
    if emo["noplan"]:
        return "暂无需求/退出阶段", "“" + emo["noplan"][0] + "”"
    if re.search(r"(确定|已定|定了).{0,8}去|已经订|订好了", t) and need["出行日期"]["date"]:
        return "已明确出行/购买阶段", "明确日期＋“确定去”"
    if emo["strong"]:
        return "决策阶段", "“" + emo["strong"][0] + "”"
    if need["目的地"].get("compare") or emo["hesitate"]:
        return "比较阶段", ("对比多个目的地" if need["目的地"].get("compare") else "“" + emo["hesitate"][0] + "”")
    if dest == "未指定":
        return "探索阶段", "目的地未指定"
    if emo["refuse"]:
        return "比较阶段", "意向未表达，且拒绝营销"
    if need["出行日期"]["date"] and t == "":
        return "决策阶段", "表单已填写目的地与日期"
    return "探索阶段" if not need["出行日期"]["date"] else "比较阶段", "仅表达目的地，未表达购买动作"


def decide_touch(emo, stage, budget_ok=True):
    """严格优先级：①拒绝营销/明确退出→不触达；②暂无计划/确定不去→不触达(暂缓)；③有兴趣但犹豫/顾虑/预算约束→谨慎；④明确准备出行→触达。"""
    if emo["refuse"]:
        return "建议不触达", "第一优先级：用户明确拒绝营销/要求不要联系（“" + emo["refuse"][0] + "”）", \
               "停止营销触达，不生成强促销内容。", "等待用户需求变化或重新授权。"
    if emo["noplan"] or stage == "暂无需求/退出阶段":
        return "建议不触达", "第二优先级：用户明确表示暂无出行计划/已确定不去（“" + (emo["noplan"][0] if emo["noplan"] else stage) + "”）", \
               "暂缓营销触达，仅保留服务类信息。", "等待用户需求变化（如再次搜索/咨询）后重新评估。"
    if emo["concerns"] or not budget_ok:
        why = "、".join(c for c, _ in emo["concerns"]) + ("、预算与目的地参考价不匹配" if not budget_ok else "")
        hi = "（但出行意向较高）" if emo["strong"] else ""
        return "谨慎触达", f"第三优先级：有出行兴趣但存在{why.strip('、')}{hi}", \
               "信息型/决策辅助型营销：价格与航班选择、替代方案、出行信息，不做强促销。", "顾虑消除或用户主动查看产品后升级为建议触达。"
    if emo["strong"] or stage in ("决策阶段", "已明确出行/购买阶段"):
        return "建议触达", "第四优先级：用户明确准备出行/订票/查看产品", "推送匹配航线与组合产品，突出时效与权益。", "—"
    return "谨慎触达", "意向信号不足（未表达明确购买动作）", "轻量种草/灵感内容，低频触达。", "出现订票/查询航班行为后升级。"


# =====================================================================
# 5. 航线/目的地匹配（目的地未指定→国内候选池优先动态推荐；已指定→分析指定目的地）
# =====================================================================
HOLIDAY_FACTOR = 1.3  # 节假日价格上浮系数（Demo参数，可替换为实时运价）


def cost_ref(dest, need):
    f = STATE["features"].get(dest)
    if not f or not f.get("ref_cost"):
        return None
    lo, hi = f["ref_cost"]
    if need["出行日期"].get("holiday"):
        lo, hi = lo * HOLIDAY_FACTOR, hi * HOLIDAY_FACTOR
    return int(lo), int(hi)


def budget_fit(dest, need):
    b = need["预算"]["value"]
    c = cost_ref(dest, need)
    if c is None or not isinstance(b, (int, float)):
        return "未知", 0.5
    lo, hi = c
    if b >= hi: return "充裕", 1.0
    if b >= (lo + hi) / 2: return "匹配", 0.9
    if b >= lo: return "偏紧", 0.6
    return "不足", 0.0


def destination_market(dest):
    """取某目的地的市场证据：候选池 → 第一轮表 → 国际扩展百度指数(xlsx) → 数据待补。"""
    p = next((x for x in STATE["pool"] if x["dest"] == dest), None)
    if p:
        return dict(src="候选池（两轮筛选）", heat=p["heat"], trend=p["trend"], rank=p["rank"], r1=p["r1"], air=p["air"], visa=p["visa"], holiday=p["holiday"], pool=True)
    r1 = STATE["round1"]
    r = r1[r1["目的地"] == dest] if not r1.empty else r1
    info = dict(src="", heat=None, trend="数据待补", rank=None, pool=False, r1=None, air=None, visa=None, holiday=None, xlsx=[])
    if len(r):
        r = r.iloc[0]
        info.update(src="第一轮粗筛表（未进入候选池）", r1=r["第一轮粗筛总分"], air=r["直飞(0/1/2)"], visa=r["签证(0/1/2)"], holiday=r["节假日适配(0/1)"])
    a = STATE["agg"].get(dest)
    if a and a.get("score") is not None:
        info.update(heat=a["score"], trend=a["trend"], src=(info["src"] + "＋第二轮关键词").strip("＋"))
    country = STATE["features"].get(dest, {}).get("country", "")
    for sn, s in STATE["xlsx"].items():
        if sn.startswith(dest) or (country and sn.startswith(country)):
            h = xlsx_heat(s)
            if h:
                info["xlsx"].append((sn, h))
    if info["xlsx"]:
        own = [h for sn, h in info["xlsx"] if sn.startswith(dest)]
        if info["heat"] is None and own:
            info.update(heat=sum(h["score"] for h in own) / len(own), trend=trend_label(sum(h["score"] for h in own) / len(own), sum(h["M"] for h in own) / len(own)))
        info["src"] = (info["src"] + "＋百度指数数据集（国际扩展验证）").strip("＋")
    if not info["src"]:
        info["src"] = "数据待补"
    return info


def match_destinations(need, emo, intl=False, stage=""):
    dep = need["出行城市"]["value"] if "出行城市" in need else need["出发城市"]["value"]
    ints = need["兴趣"]["value"] if isinstance(need["兴趣"]["value"], list) else []
    n = need["出行人数"]["value"] if isinstance(need["出行人数"]["value"], int) else None
    pool = [p for p in STATE["pool"] if p["dest"] != dep]
    dom = [p for p in pool if dest_type(p["dest"]) == "国内"]
    cand = (dom + [p for p in pool if p not in dom]) if intl else dom
    hmax = max((p["heat"] or 0) for p in STATE["pool"]) or 1
    res = []
    for p in cand:
        d = p["dest"]
        f = STATE["features"].get(d, {})
        tags = f.get("tags", [])
        overlap = [i for i in ints if i in tags or (i == "家庭" and "亲子" in tags)]
        isc = len(overlap) / len(ints) if ints else 0.5
        bl, bsc = budget_fit(d, need)
        hol = need["出行日期"].get("holiday")
        feas = ((p["air"] or 0) / 2 + (p["visa"] or 0) / 2 + (p["holiday"] or 0)) / 3
        party = 1.0 if not n or n <= 2 or "亲子" in tags else 0.7
        score = 0.35 * (p["heat"] or 0) / hmax + 0.25 * isc + 0.25 * bsc + 0.10 * feas + 0.05 * party
        if dest_type(d) != "国内":
            score -= 0.05  # 国内为主：国际仅作扩展
        reasons = [f"市场热度：第二轮综合热度分 {fmt(p['heat'])}（候选池第{p['rank']}名）",
                   f"市场趋势：{p['trend']}",
                   f"用户兴趣：{'匹配 ' + '、'.join(overlap) if overlap else ('未提供兴趣，按热度推荐' if not ints else '兴趣匹配度低')}（目的地特色：{f.get('highlights', '—')}）",
                   f"用户预算：{bl}" + (f"（参考 {cost_ref(d, need)[0]}–{cost_ref(d, need)[1]} 元/人{'，已计节假日上浮' if hol else ''}）" if cost_ref(d, need) else ""),
                   f"出行日期：{'节假日适配度高' if (p['holiday'] or 0) >= 1 and hol else '节假日适配' if (p['holiday'] or 0) >= 1 else '节假日适配数据待补'}",
                   f"航线可行性：航空{fmt(p['air'],0)}/2、签证/入境{fmt(p['visa'],0)}/2（数据表口径以上海出发为基准，其他出发城市仅作参考）",
                   f"人数：{n or '未指定'}人" + ("，适合同行" if party == 1.0 else "，多人出行需关注行程统一"),
                   f"决策阶段：{stage or '未识别'}（决定触达强度与内容类型）",
                   f"出发城市约束：{dep}→{d}（排除与出发地相同的目的地）；百度指数为目的地市场级信号，不要求限定出发地"]
        res.append(dict(dest=d, score=score, budget=bl, overlap=overlap, reasons=reasons, type=dest_type(d),
                        air=p["air"], holiday=p["holiday"], heat=p["heat"], trend=p["trend"], rank=p["rank"]))
    res.sort(key=lambda x: -x["score"])
    ok = [r for r in res if r["budget"] != "不足"]
    return ok, [r for r in res if r["budget"] == "不足"]


# =====================================================================
# 6. 动态旅客画像
# =====================================================================
def build_profile(need, emo, stage, decision, focus):
    b = need["预算"]["value"]
    n = need["出行人数"]["value"]
    ints = need["兴趣"]["value"] if isinstance(need["兴趣"]["value"], list) else []
    bl = "未知预算" if not isinstance(b, (int, float)) else "预算敏感型" if b < 3000 else "性价比型" if b < 5000 else "中高消费型" if b < 8000 else "品质型"
    if emo["price"] and bl not in ("预算敏感型",):
        bl = "价格敏感型"
    party = {1: "单人", 2: "双人"}.get(n, "家庭/小团" if isinstance(n, int) and n <= 5 else "团体" if isinstance(n, int) else "")
    if "亲子" in ints or "家庭" in ints:
        party = "亲子家庭"
    focus_t = "".join(ints[:2]) + "导向型" if ints else "综合型"
    mode = "商旅" if "商务" in ints else "自由行"
    stage_t = {"探索阶段": "探索期", "比较阶段": "犹豫比较型", "决策阶段": "高意向", "已明确出行/购买阶段": "已明确出行", "暂无需求/退出阶段": "暂无需求"}[stage]
    hol = need["出行日期"].get("holiday")
    name = f"{stage_t}·{bl}{focus_t}{party}{mode}旅客" + ("（节假日出行）" if hol else "")
    if decision == "建议不触达":
        name = f"{stage_t}·{'拒绝营销' if emo['refuse'] else '暂无出行需求'}用户（仅保留服务关系）"
    d = need["出行日期"].get("date")
    days = (d - dt.date.today()).days if d else None
    dims = [("预算层级", bl, f"预算 {b if b != '未指定' else '未指定'}"),
            ("同行结构", party or "未指定", f"出行人数 {n}"),
            ("兴趣偏好", "、".join(ints) or "未指定", need["兴趣"]["evidence"]),
            ("出行时间", need["出行日期"]["value"] + (f"（距今{days}天）" if days is not None else ""), need["出行日期"]["evidence"]),
            ("决策阶段", stage, "见决策阶段识别"),
            ("情绪/意向", emo["mood"], "、".join(c for c, _ in emo["concerns"]) or "无明显顾虑"),
            ("关注目的地", focus or "待推荐", "用户指定" if need["目的地"]["value"] != "未指定" else "系统基于候选池推荐"),
            ("偏好/排除", "、".join((need["偏好"]["value"] if isinstance(need["偏好"]["value"], list) else []) + (need["排除条件"]["value"] if isinstance(need["排除条件"]["value"], list) else [])) or "未提及", "原文提取")]
    m = destination_market(focus) if focus else None
    if m and m["heat"] is not None:
        dims.append(("市场信号", f"{focus}热度 {fmt(m['heat'])}，{m['trend']}", m["src"]))
    content = []
    if "价格" in bl or "预算敏感" in bl: content += ["价格信息与低价日历", "高性价比航班选择"]
    if emo["visa"]: content.append("签证/入境办理指引")
    if stage == "探索阶段": content.append("目的地灵感与行程模板")
    if stage == "比较阶段": content.append("目的地/航班对比与决策辅助")
    if stage in ("决策阶段", "已明确出行/购买阶段"): content.append("航班+酒店组合与权益")
    content += [f"{i}主题内容" for i in ints[:2]]
    return dict(name=name, dims=dims, content=content, days=days)


# =====================================================================
# 7. 渠道与时机 / AIGC / 人工审核
# =====================================================================
def channel_timing(decision, stage, days, emo):
    if decision == "建议不触达":
        return "不触达（停止营销渠道推送）", "不安排触达；仅在用户主动咨询时提供服务"
    ch = {"探索阶段": "小红书/抖音内容种草 + 东航App灵感频道", "比较阶段": "东航App比价/低价日历 + 微信公众号攻略",
          "决策阶段": "东航App推送 + 短信（已授权）", "已明确出行/购买阶段": "东航App + 短信：航班/酒店/值机服务提醒"}.get(stage, "东航App")
    if decision == "谨慎触达":
        ch = ch.split("+")[0].strip() + "（低频、信息型）"
    if days is None:
        tm = "用户活跃时段（晚间19:00–22:00）低频触达；补充出行日期后再精准安排"
    elif days <= 3:
        tm = "立即（出行临近，24小时内推送服务与补充产品）"
    elif days <= 14:
        tm = f"48小时内触达（距出行{days}天，决策窗口期）"
    else:
        tm = f"出行前约{min(days, 21)}天起，每周不超过1次（距出行{days}天）"
    if emo["conv"] and any("早" in c for c in emo["conv"]):
        tm += "；推荐航班避开早班"
    return ch, tm


def ch_hint(decision, stage):
    return "东航App/小程序站内推送与订阅提醒（低打扰）" if decision == "谨慎触达" else "东航App推送＋短信/私域社群＋重点航线信息流"


def aigc_content(need, prof, decision, stage, emo, recs, focus):
    if decision == "建议不触达":
        return "（建议不触达：不生成营销文案。）"
    dep = need["出发城市"]["value"] if need["出发城市"]["value"] != "未指定" else "您的出发地"
    n = need["出行人数"]["value"]
    nt = f"{n}人" if isinstance(n, int) else ""
    hol = need["出行日期"].get("holiday") or need["出行日期"]["value"].split("（")[0]
    ints = need["兴趣"]["value"] if isinstance(need["兴趣"]["value"], list) else []
    f = STATE["features"].get(focus, {}) if focus else {}
    c = cost_ref(focus, need) if focus else None
    price_line = f"参考人均 {c[0]}–{c[1]} 元（机票+酒店估算，非实时价，以下单页为准）" if c else "价格以东航App实时查询为准"
    excl = need["排除条件"]["value"] if isinstance(need["排除条件"]["value"], list) else []
    prefs = need["偏好"]["value"] if isinstance(need["偏好"]["value"], list) else []
    lines = []
    if decision == "谨慎触达" or "价格" in prof["name"] or "预算敏感" in prof["name"]:
        lines.append(f"【标题】{dep}→{focus}：{nt}{hol}出行，怎么飞更省心省钱？")
        lines.append(f"· 价格信息：{price_line}；可开启低价提醒，价格变化第一时间告知。")
        lines.append("· 航班选择：对比直飞/中转与不同时段票价" + ("，已为您过滤过早的航班" if any("早" in e for e in excl) else "") + "。")
        if emo["visa"] or dest_type(focus) != "国内":
            lines.append("· 出行信息：签证/入境材料清单与办理时长说明。")
        lines.append(f"· 套餐价值：机票+{'交通便利' if '交通方便' in prefs else '精选'}酒店组合，比单订更透明。")
        if recs:
            lines.append("· 决策辅助：同预算还可考虑 " + "、".join(r["dest"] for r in recs[:2] if r["dest"] != focus) + "。")
        lines.append("（不催促下单，您可以先收藏比较。）")
    else:
        lines.append(f"【标题】{nt}{hol}，从{dep}飞{focus}，{('、'.join(ints[:2]) + '之旅') if ints else '说走就走'}！")
        lines.append(f"· 推荐理由：{f.get('highlights', focus)}，与您的{('、'.join(ints) or '出行')}偏好匹配。")
        lines.append(f"· 航班+酒店：{price_line}" + ("；已优先展示非早班航班" if any("早" in e for e in excl) else "") + ("；酒店优选交通便利地段" if "交通方便" in prefs else "") + "。")
        if stage == "已明确出行/购买阶段":
            lines.append("· 出行服务：在线值机、行李额度与目的地交通提醒。")
        lines.append("· 会员权益：东航会员积分与升级权益（以实际活动为准）。")
    body = "\n".join(lines)
    route = f"{dep}→{focus}" if focus else "—"
    visual = (f"海报建议：以“{'、'.join(ints[:2]) if ints else focus}”为主视觉，突出{route}与" + ("价格提醒/低价日历" if decision == "谨慎触达" else "出行服务与限时权益") + "；配色沿用东航品牌色，文字留白用于价格与日期。")
    head = (f"【客群】{prof['name']}\n【推荐航线】{route}\n"
            f"【推荐理由】" + (recs[0]["reasons"][0] if recs and recs[0]["dest"] == focus else f"围绕用户需求（{'、'.join(ints) or '未提供兴趣'}）与市场趋势")
            + f"\n【营销渠道】{ch_hint(decision, stage)}\n")
    return head + "\n【营销文案】\n" + body + f"\n\n【视觉/海报建议】{visual}\n\n⚠ AIGC生成，需人工审核。{AIGC_NOTE}｜人工审核状态：待审核"


# =====================================================================
# 8. 主流程：智能营销
# =====================================================================
TASK_MIN, TASK_MAX = 3, 5


def build_task(need, recs, specified, focus, decision):
    """营销任务 = 一个出发城市 + 一个假期/出行时间 + 3～5条候选航线。
    目的地只来自最终候选目的地池（两轮筛选结果），不随机生成、不补齐；指定目的地时直接分析该航线。"""
    dep = need["出发城市"]["value"]
    hd = need["出行日期"]
    hol = hd.get("holiday")
    label = hol or (hd["value"] if hd["value"] != "未指定" else "未指定")
    T = dict(city=dep, holiday=label, specified=specified, routes=[], dropped=[], eligible=0, insufficient=False, note="")
    if specified:
        T["routes"] = [dict(dest=focus)]
        T["note"] = f"用户已指定航线：{dep}→{focus}，直接分析该航线，不替换成其他目的地。"
        return T
    if decision == "建议不触达":
        T["note"] = "触达决策为【建议不触达】：不生成候选航线与营销方案。"
        return T
    elig = []
    for r in recs:
        if hol and r.get("holiday") == 0:
            T["dropped"].append((r["dest"], f"{hol}节假日适配不足"))
        elif r.get("air") == 0:
            T["dropped"].append((r["dest"], "暂无可行航线"))
        else:
            elig.append(r)
    T["eligible"] = len(elig)
    T["routes"] = elig[:TASK_MAX]
    n = len(T["routes"])
    if n < TASK_MIN:
        T["insufficient"] = True
        T["note"] = f"候选不足：当前符合条件的目的地仅 {n} 个（要求 {TASK_MIN}～{TASK_MAX} 条），系统不会随机补齐，请放宽预算/假期条件或更新市场数据。"
    elif len(elig) > TASK_MAX:
        T["note"] = f"共 {len(elig)} 个符合条件的目的地，按热度、趋势、预算、兴趣、航线可行性、假期适配综合得分保留前 {TASK_MAX} 条。"
    else:
        T["note"] = f"共 {n} 条候选航线，全部来自当前最终候选目的地池。"
    return T


def task_html(T):
    if not T:
        return ""
    why = ""
    if T["dropped"]:
        why = '<br><span style="font-size:12px;color:#6b7280">已排除：' + "、".join(f"{esc(d)}（{esc(w)}）" for d, w in T["dropped"]) + "</span>"
    col = "#dc2626" if T["insufficient"] else "#16a34a"
    return (f'<div style="{CARD};max-width:none"><b>营销任务</b>：出发城市 <b>{esc(str(T["city"]))}</b> ＋ 假期/出行时间 <b>{esc(str(T["holiday"]))}</b> ＋ '
            f'候选航线 <b>{len(T["routes"])}</b> 条（范围 {TASK_MIN}～{TASK_MAX}）<br>'
            f'<span style="color:{col}">{esc(T["note"])}</span>{why}'
            f'<br><span style="font-size:12px;color:#6b7280">群体信号＝最终候选目的地池（市场热度/趋势，百度指数为市场级信号）；个体信号＝用户输入（预算/人数/兴趣/出发城市/假期/决策阶段）。出发城市作为航线匹配约束；航线可行性数据以数据表“上海出发”口径为基准。</span></div>')


def run_pipeline(text, city, dest, budget, n, tags, date, intl, holiday="", need=None, src=None):
    if not STATE["pool"]:
        refresh_market_state()
    form = dict(city=city, dest=dest, budget=budget, n=n, tags=tags or [], date=date, holiday=holiday if holiday in HOLIDAYS else "")
    need = need or analyze_user_need(text, form)
    if intl and need["目的地"]["value"] == "未指定":
        need["目的地"]["intl"] = True
    emo = analyze_emotion(text, need)
    stage, stage_ev = detect_stage(text, need, emo)
    dval = need["目的地"]["value"]
    specified = dval != "未指定"
    focus = dval.split("、")[0] if specified else None
    recs, excluded = match_destinations(need, emo, intl=bool(intl) or (specified and dest_type(focus) != "国内"), stage=stage)
    budget_ok = True
    alt = []
    if specified:
        bl, _ = budget_fit(focus, need)
        budget_ok = bl not in ("不足",)
        if bl in ("不足", "偏紧"):
            alt = [r for r in recs if r["budget"] in ("充裕", "匹配") and r["type"] == "国内"][:3]
    else:
        if recs:
            focus = recs[0]["dest"]
        else:
            budget_ok = False
    decision, reason, strategy, reeval = decide_touch(emo, stage, budget_ok)
    if decision == "建议不触达" and not specified:
        focus, recs = None, []  # 不触达且未指定目的地：不做营销推荐
    task = build_task(need, recs, specified, focus, decision)
    if not specified and decision != "建议不触达" and task["routes"]:
        focus = task["routes"][0]["dest"]  # 文案围绕任务内排名第一的航线
    prof = build_profile(need, emo, stage, decision, focus)
    ch, tm = channel_timing(decision, stage, prof["days"], emo)
    content = aigc_content(need, prof, decision, stage, emo, recs, focus) if focus else (
        "（当前预算低于所有候选目的地参考价，不生成促销文案；建议推送“低价日历/错峰出行”信息型内容。）\n\n⚠ " + AIGC_NOTE if decision != "建议不触达" else "（建议不触达：不生成营销文案。）")
    R = dict(source=src or dict(kind="内部手动录入（测试/电话咨询，非客户提交）", id="—", time="—"), task=task, need=need, emo=emo, stage=stage, stage_ev=stage_ev, specified=specified, focus=focus, recs=recs, excluded=excluded, alt=alt,
             decision=decision, reason=reason, strategy=strategy, reeval=reeval, profile=prof, channel=ch, timing=tm, content=content,
             budget_ok=budget_ok, time=dt.datetime.now().strftime("%H:%M:%S"), market_time=STATE["time"])
    return R


def layers_html(R):
    """内部决策台：明确区分群体层（市场信号）与个体层（客户需求）。"""
    src = R.get("source") or {}
    n = len(STATE["pool"])
    return (f'<div style="display:flex;gap:10px;flex-wrap:wrap;margin:6px 0">'
            f'<div style="{CARD}"><b>群体层 · 市场信号</b><br><span style="font-size:12px">来源：平台榜单/热词·媒体社交（第一轮粗筛依据）＋百度指数（第二轮精筛）；'
            f'决策台只读取“最终候选目的地池”（{n}个）；市场数据版本 {esc(str(R.get("market_time", "")))}。回答：市场整体有什么需求。</span></div>'
            f'<div style="{CARD}"><b>个体层 · 客户需求</b><br><span style="font-size:12px">来源：{esc(str(src.get("kind", "")))}'
            f'（编号 {esc(str(src.get("id", "—")))}，需求生成时间 {esc(str(src.get("time", "—")))}）。回答：这个客户有什么需求。</span></div></div>'
            f'<p style="font-size:12px;color:#6b7280">群体层市场信号 ＋ 个体层客户需求 → 动态旅客画像 → 候选航线 → 触达决策</p>')


def need_table(need):
    rows = ""
    for k in ["出发城市", "目的地", "出行人数", "预算", "出行日期", "兴趣", "偏好", "排除条件"]:
        f = need[k]
        v = f["value"]
        v = "、".join(v) if isinstance(v, list) else (f"{v}{need['预算'].get('per', '')}" if k == "预算" and isinstance(v, int) else (f"{v}人" if k == "出行人数" and isinstance(v, int) else v))
        col = {"高": "#16a34a", "中": "#ca8a04", "低": "#dc2626"}[f["conf"]]
        rows += f'<tr><td {TD}>{k}</td><td {TD}><b>{esc(str(v))}</b></td><td {TD}>{esc(f["evidence"])}</td><td {TD} style="color:{col}">{f["conf"]}</td></tr>'
    tip = '<p style="color:#dc2626;font-size:13px">目的地未指定（置信度低）：系统不擅自猜测目的地，将从国内候选池动态推荐，请用户确认或修改。</p>' if need["目的地"]["value"] == "未指定" else ""
    return f'<table style="border-collapse:collapse;font-size:13px"><tr><th {TH}>字段</th><th {TH}>提取结果</th><th {TH}>原文依据</th><th {TH}>置信度</th></tr>{rows}</table>{tip}'


def result_html(R):
    need, emo = R["need"], R["emo"]
    dcol = {"建议触达": "#16a34a", "谨慎触达": "#ca8a04", "建议不触达": "#dc2626"}[R["decision"]]
    # 推荐/分析
    if R["specified"]:
        m = destination_market(R["focus"])
        bl, _ = budget_fit(R["focus"], need)
        c = cost_ref(R["focus"], need)
        xl = "".join(f"<li>{esc(sn)}：最新 {h['latest']}（{h['date']:%Y-%m-%d}），综合热度 {h['score']:.1f}，{h['trend']}</li>" for sn, h in m.get("xlsx", []))
        rec = (f'<p><b>目的地已指定 → 分析用户指定目的地：{esc(R["focus"])}</b>（{dest_type(R["focus"])}{"，国际扩展场景" if dest_type(R["focus"]) != "国内" else ""}）</p><ul>'
               f'<li>市场证据来源：{m["src"]}</li><li>市场热度：{fmt(m["heat"])}；趋势：{m["trend"]}' + (f'；候选池第{m["rank"]}名' if m.get("rank") else "；不在当前候选池（不会因此替换用户目的地）") + '</li>'
               f'<li>航线/入境可行性：航空 {fmt(m["air"],0)}/2，签证 {fmt(m["visa"],0)}/2' + ("" if m["air"] is not None else "（数据待补）") + '</li>'
               f'<li>预算匹配：{bl}' + (f'（参考 {c[0]}–{c[1]} 元/人）' if c else "（无参考价，数据待补）") + f'</li>{xl}</ul>')
        if R["alt"]:
            rec += "<p><b>预算约束下的国内备选（供用户参考，不替换指定目的地）：</b>" + "、".join(f'{a["dest"]}（{a["budget"]}）' for a in R["alt"]) + "</p>"
    else:
        T = R["task"]
        dep = need["出发城市"]["value"]
        if T["routes"]:
            items = "".join(f'<li><b>{i+1}. {esc(str(dep))}→{esc(r["dest"])}</b>（{r["type"]}，匹配分 {r["score"]*100:.0f}）<ul>' + "".join(f"<li>{esc(x)}</li>" for x in r["reasons"]) + "</ul></li>"
                            for i, r in enumerate(T["routes"]))
            rec = f'<p><b>目的地未指定 → 从最终候选目的地池动态生成候选航线</b>（市场热度＋趋势＋预算＋日期＋人数＋兴趣＋航线可行性＋决策阶段，非随机）</p><ol style="list-style:none;padding-left:0">{items}</ol>'
        else:
            rec = "<p><b>当前预算低于全部候选目的地的参考价</b>，不输出不匹配的营销方案。建议：提高预算、错峰出行或选择周边短途。</p>"
        if R["excluded"]:
            rec += f'<p style="font-size:12px;color:#6b7280">因预算不足未推荐：{"、".join(e["dest"] for e in R["excluded"])}</p>'
    conc = "、".join(f"{c}（{e}）" for c, e in emo["concerns"]) or "无"
    return (f'<p style="font-size:12px;color:#6b7280">分析时间 {R["time"]}｜使用市场数据版本：{R["market_time"]}</p>'
            f'{layers_html(R)}'
            f'<h3>1. 需求识别</h3>{need_table(need)}'
            f'<h3>2. 旅客画像</h3><p><b>{esc(R["profile"]["name"])}</b>（详见“旅客画像”页）</p>'
            f'<h3>3~4. 营销任务、候选航线与推荐理由</h3>{task_html(R["task"])}{rec}'
            f'<h3>5. 决策阶段</h3><p>{R["stage"]}（依据：{esc(R["stage_ev"])}）</p>'
            f'<h3>6. 情绪/顾虑</h3><p>出行意愿：{emo["intent"]}；情绪倾向：{emo["mood"]}；顾虑：{esc(conc)}</p>'
            f'<h3>7. 触达决策</h3><div style="{CARD};max-width:none;border-color:{dcol}"><span style="font-size:20px;font-weight:700;color:{dcol}">【{R["decision"]}】</span><br>'
            f'原因：{esc(R["reason"])}<br>营销策略：{esc(R["strategy"])}<br>重新评估条件：{esc(R["reeval"])}</div>'
            f'<h3>8. 营销渠道</h3><p>{esc(R["channel"])}</p><h3>9. 最佳触达时机</h3><p>{esc(R["timing"])}</p>')


def profile_html(R, prev=None):
    if not R:
        return "<p>请先在“智能营销”页输入用户需求并提交，系统将根据需求提取结果与市场数据动态生成旅客画像。</p>"
    p = R["profile"]
    rows = "".join(f'<tr><td {TD}>{k}</td><td {TD}><b>{esc(str(v))}</b></td><td {TD}>{esc(str(e))}</td></tr>' for k, v, e in p["dims"])
    diff = ""
    if prev:
        pp = prev["profile"]
        ch = [(k, ov, nv) for (k, ov, _), (k2, nv, _) in zip(pp["dims"], p["dims"]) if k == k2 and ov != nv]
        recs_o = "、".join(r["dest"] for r in prev["recs"][:3]) or "无"
        recs_n = "、".join(r["dest"] for r in R["recs"][:3]) or "无"
        diff = (f'<h3>与上一版本对比（动态更新证明）</h3><table style="border-collapse:collapse;font-size:13px"><tr><th {TH}>项目</th><th {TH}>上一版本</th><th {TH}>当前版本</th></tr>'
                f'<tr><td {TD}>画像</td><td {TD}>{esc(pp["name"])}</td><td {TD}><b>{esc(p["name"])}</b></td></tr>'
                + "".join(f'<tr><td {TD}>{k}</td><td {TD}>{esc(str(o))}</td><td {TD}><b>{esc(str(nv))}</b></td></tr>' for k, o, nv in ch) +
                f'<tr><td {TD}>推荐/分析目的地</td><td {TD}>{esc(prev["focus"] or "无")}｜候选 {esc(recs_o)}</td><td {TD}><b>{esc(R["focus"] or "无")}</b>｜候选 {esc(recs_n)}</td></tr>'
                f'<tr><td {TD}>触达决策</td><td {TD}>{prev["decision"]}</td><td {TD}><b>{R["decision"]}</b></td></tr>'
                f'<tr><td {TD}>营销渠道</td><td {TD}>{esc(prev["channel"])}</td><td {TD}><b>{esc(R["channel"])}</b></td></tr></table>')
    fbh = ""
    if R.get("feedback"):
        fbh = f'<h3>效果回填后的画像与策略更新</h3><p style="color:#dc2626;font-size:12px">{SIM_NOTE}</p><ul>' + "".join(f"<li>{esc(x)}</li>" for x in R["feedback"]) + "</ul>"
    return (layers_html(R) + f'<h3>动态旅客画像：<span style="color:#f97316">{esc(p["name"])}</span></h3>'
            f'<p style="font-size:13px">画像由“用户输入＋需求提取结果＋市场数据＋兴趣＋预算＋人数＋出行日期＋决策阶段＋情绪/意向”实时生成，非固定模板；输入改变后画像随之变化。</p>'
            f'<table style="border-collapse:collapse;font-size:13px"><tr><th {TH}>画像维度</th><th {TH}>结果</th><th {TH}>依据</th></tr>{rows}</table>'
            f'<p><b>内容偏好（用于AIGC）：</b>{esc("、".join(p["content"]) or "—")}</p>{fbh}{diff}')


# =====================================================================
# 9. 效果评估（Demo模拟数据）与数据反馈闭环
# =====================================================================
def simulate_effect(R, review):
    if not R:
        return None
    if R["decision"] == "建议不触达" or review != "审核通过":
        return dict(skip=True)
    seed = int(hashlib.md5((R["profile"]["name"] + R["decision"] + str(R["focus"])).encode()).hexdigest()[:8], 16)
    base = 20000 if R["decision"] == "建议触达" else 8000
    exp = base + seed % 5000
    ctr = (0.06 if R["decision"] == "建议触达" else 0.04) + (seed % 30) / 1000
    click = int(exp * ctr)
    cons = int(click * (0.22 + (seed % 10) / 100))
    buy = int(cons * (0.30 if R["stage"] in ("决策阶段", "已明确出行/购买阶段") else 0.15))
    c = cost_ref(R["focus"], R["need"]) if R["focus"] else None
    price = int(sum(c) / 2) if c else 3000
    cost = exp * 0.12
    return dict(skip=False, exp=exp, click=click, cons=cons, buy=buy, price=price, cost=cost, ctr=click / exp, cvr=buy / max(click, 1),
                roi=(buy * price * 0.08 - cost) / cost)  # ROI按航司毛利8%估算（模拟口径）


def feedback_items(E):
    """效果回填 → 重新分析用户 → 更新画像 → 优化营销策略（规则化，基于Demo模拟指标）。"""
    fb = []
    if E["ctr"] < 0.07: fb.append("点击率偏低 → 重新分析用户：内容与兴趣匹配不足 → 更新画像内容偏好 → 下一轮突出用户兴趣主题与目的地亮点")
    if E["cvr"] < 0.05: fb.append("转化率偏低 → 用户仍在比较/存在顾虑 → 画像标注“犹豫比较” → 下一轮改为价格日历＋决策辅助信息型内容")
    if E["roi"] < 0: fb.append("ROI为负 → 收缩高成本渠道，保留东航App/私域低成本触达")
    if not fb: fb.append("各项指标良好 → 保持当前策略，扩大相似画像人群投放")
    return fb


def effect_html(R, review):
    head = f'<p style="color:#dc2626;font-weight:700">{SIM_NOTE}</p>'
    if not R:
        return head + "<p>请先在“智能营销”页生成方案并完成人工审核。</p>", None
    E = simulate_effect(R, review)
    if E["skip"]:
        why = "触达决策为【建议不触达】，未投放" if R["decision"] == "建议不触达" else f"人工审核状态为【{review}】，内容未上线"
        return head + f"<p>{why}，暂无效果数据。</p><h3>数据反馈</h3><p>保持当前画像，等待用户行为变化或审核通过后再评估。</p>", None
    cards = "".join(f'<div style="{CARD}"><div style="font-size:12px;color:#6b7280">{k}（模拟）</div><div style="font-size:24px;font-weight:700">{v}</div></div>'
                    for k, v in [("曝光", f'{E["exp"]:,}'), ("点击", f'{E["click"]:,}'), ("咨询", f'{E["cons"]:,}'), ("购票", f'{E["buy"]:,}'),
                                 ("点击率", f'{E["ctr"]*100:.1f}%'), ("转化率", f'{E["cvr"]*100:.1f}%'), ("ROI", f'{E["roi"]:.2f}')])
    fb = feedback_items(E)
    flow = "营销效果 → 重新分析用户 → 更新旅客画像 → 优化营销策略 → 进入下一轮"
    return (head + f'<p>评估对象：{esc(R["profile"]["name"])}｜目的地：{esc(R["focus"] or "—")}｜决策：{R["decision"]}</p>'
            f'<div style="display:flex;gap:10px;flex-wrap:wrap">{cards}</div>'
            f'<p style="font-size:12px;color:#6b7280">模拟口径：客单价取目的地参考价中值 {E["price"]} 元；营销成本=曝光×0.12元；ROI=(购票×客单价×8%毛利−成本)÷成本。</p>'
            f'<h3>数据反馈 → 下一轮营销优化</h3><p><b>{flow}</b></p><ul>' + "".join(f"<li>{esc(x)}</li>" for x in fb) + "</ul>"), effect_fig(E)


def effect_fig(E):
    try:
        import plotly.graph_objects as go
        weeks = [f"第{i}周" for i in range(1, 5)]
        w = [0.18, 0.24, 0.28, 0.30]
        fig = go.Figure()
        for i, (nm, k) in enumerate([("曝光", "exp"), ("点击", "click"), ("咨询", "cons"), ("购票", "buy")]):
            fig.add_scatter(x=weeks, y=[max(1, int(E[k] * x)) for x in w], name=nm, mode="lines+markers", line=dict(color=COLORS[i]))
        fig.update_layout(template="plotly_white", height=360, yaxis=dict(type="log", title="人数（Demo模拟，对数坐标）"), title="营销趋势（Demo模拟）",
                          margin=dict(l=50, r=30, t=50, b=40), legend=dict(orientation="h", y=1.15))
        return fig
    except ImportError:
        return None


# =====================================================================
# 9b. 客户侧 → 内部侧：结构化需求对象的生成、存储与转换（真实数据传递）
# =====================================================================
REQ_FILE = os.path.join(DATA_DIR, "customer_requests.json")  # 运行时生成；客户提交的需求在此保存，内部决策台读取
REQ_FIELDS = ["出发城市", "出行时间", "出行人数", "预算", "目的地", "兴趣", "偏好", "排除条件"]
_REQ_MEM = []


def _ensure():
    if not STATE["pool"]:
        refresh_market_state()


def load_requests():
    try:
        with open(REQ_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return list(_REQ_MEM)


def save_request(req):
    reqs = load_requests()
    req["id"] = f"REQ-{dt.datetime.now():%Y%m%d%H%M%S}-{len(reqs) + 1:03d}"
    reqs.append(req)
    _REQ_MEM[:] = reqs
    try:
        os.makedirs(os.path.dirname(REQ_FILE), exist_ok=True)
        with open(REQ_FILE, "w", encoding="utf-8") as f:
            json.dump(reqs, f, ensure_ascii=False, indent=1)
    except Exception:
        pass  # 文件不可写时仍保留在内存，同一进程内决策台可读取
    return req["id"]


def _blank(v):
    return None if v in (None, "", "未指定", "未提及") else v


def _split(x):
    if isinstance(x, list):
        return [i for i in x if i]
    return [i.strip() for i in re.split(r"[、,，;；\s]+", str(x or "")) if i.strip()]


def customer_extract(text, city="", time_="", n=None, bud=None, dest="", tags=None, prefs="", excl=""):
    """客户侧：自然语言/表单 → 结构化字段（含原文依据、置信度）。只返回提取结果，不含任何内部营销判断。"""
    _ensure()
    form = dict(city=_blank(city), dest=_blank(dest) or "", budget=_blank(bud), n=_blank(n), tags=list(tags or []))
    t = str(time_ or "").strip()
    hol = next((h for h in HOLIDAYS if h in t), None)
    if hol:
        form["holiday"] = hol
    elif re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", t):
        form["date"] = t
    need = analyze_user_need(text, form)
    extra = " ".join(x for x in (str(prefs or ""), str(excl or "")) if x)
    if extra and not text:
        need = analyze_user_need(extra, form)
    return need


def need_to_fields(need):
    """提取结果 → 客户可编辑字段值（出行时间优先给假期名）。"""
    d = need["出行日期"]
    time_ = d.get("holiday") or (str(d["value"]).split("（")[0] if d["value"] != "未指定" else "")
    g = lambda k: _blank(need[k]["value"])
    lst = lambda k: need[k]["value"] if isinstance(need[k]["value"], list) else []
    return dict(city=g("出发城市") or "", time=time_, n=g("出行人数"), bud=g("预算"), dest=g("目的地") or "未指定",
                tags=[x for x in lst("兴趣") if x in TAGS], prefs="、".join(lst("偏好")), excl="、".join(lst("排除条件")))


def build_request(text, city, time_, n, bud, dest, tags, prefs, excl, need=None):
    """客户确认后生成结构化需求对象（至少含：出发城市、出行时间、出行人数、预算、目的地、兴趣、偏好、排除条件、原始自然语言、字段置信度、需求生成时间）。"""
    need = need or customer_extract(text, city, time_, n, bud, dest, tags, prefs, excl)
    key = {"出发城市": "出发城市", "出行时间": "出行日期", "出行人数": "出行人数", "预算": "预算", "目的地": "目的地", "兴趣": "兴趣", "偏好": "偏好", "排除条件": "排除条件"}
    req = {
        "出发城市": _blank(city) or "未指定",
        "出行时间": _blank(time_) or "未指定",
        "出行人数": int(n) if _blank(n) else "未指定",
        "预算": int(bud) if _blank(bud) else "未指定",
        "预算单位": need["预算"].get("per", "元/人"),
        "目的地": _blank(dest) or "未指定",
        "兴趣": _split(tags),
        "偏好": _split(prefs),
        "排除条件": _split(excl),
        "原始自然语言": str(text or "").strip(),
        "字段置信度": {k: need[v]["conf"] for k, v in key.items()},
        "原文依据": {k: need[v]["evidence"] for k, v in key.items()},
        "客户已确认": True,
        "需求生成时间": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return req


def need_from_request(req):
    """内部侧：把客户确认后的结构化需求还原为分析用 need 对象（以客户确认值为准，不再使用任何固定测试数据）。"""
    _ensure()
    dest = _blank(req.get("目的地"))
    need = analyze_user_need("", dict(city=_blank(req.get("出发城市")), dest=dest or "", budget=_blank(req.get("预算")),
                                      n=_blank(req.get("出行人数")), tags=list(req.get("兴趣") or [])))
    tm = _blank(req.get("出行时间"))
    if tm:
        need["出行日期"] = analyze_user_need(str(tm))["出行日期"]
    for k, v in (("偏好", req.get("偏好")), ("排除条件", req.get("排除条件"))):
        need[k] = dict(value=list(v) if v else "未提及", evidence="客户确认", conf="中")
    need["_text"] = req.get("原始自然语言", "")
    key = {"出发城市": "出发城市", "出行时间": "出行日期", "出行人数": "出行人数", "预算": "预算", "目的地": "目的地", "兴趣": "兴趣", "偏好": "偏好", "排除条件": "排除条件"}
    for k, nk in key.items():
        ev = (req.get("原文依据") or {}).get(k)
        cf = (req.get("字段置信度") or {}).get(k)
        if nk in need:
            if ev and ev not in ("结构化表单", "—"):
                need[nk]["evidence"] = f"{ev}（客户已确认）"
            elif need[nk].get("value") not in ("未指定", "未提及"):
                need[nk]["evidence"] = "客户确认填写"
            if cf in ("高", "中", "低"):
                need[nk]["conf"] = cf
    if req.get("预算单位"):
        need["预算"]["per"] = req["预算单位"]
    return need


def run_from_request(req):
    need = need_from_request(req)
    n = need["出行人数"]["value"]
    return run_pipeline(req.get("原始自然语言", ""), need["出发城市"]["value"], need["目的地"]["value"], need["预算"]["value"] if isinstance(need["预算"]["value"], (int, float)) else None,
                        n if isinstance(n, int) else None, req.get("兴趣") or [], "", False, "", need=need,
                        src=dict(kind="客户侧“结伴出游模拟器”提交（真实结构化需求）", id=req.get("id", "—"), time=req.get("需求生成时间", "—")))


def request_html(req):
    if not req:
        return "<p>暂无客户提交的需求。请客户在“结伴出游模拟器”中提交后点击【刷新客户需求】。</p>"
    need = need_from_request(req)
    conf = "、".join(f"{k}:{v}" for k, v in (req.get("字段置信度") or {}).items())
    return (f'<p><b>编号：</b>{esc(req.get("id", "—"))}　<b>需求生成时间：</b>{esc(req.get("需求生成时间", "—"))}　<b>来源：</b>客户侧提交（个体层）</p>'
            f'<p><b>原始自然语言：</b>{esc(req.get("原始自然语言", "") or "（客户未填写自然语言，仅提交结构化字段）")}</p>'
            f'{need_table(need)}<p style="font-size:12px;color:#6b7280">字段置信度：{esc(conf)}</p>')


def customer_table_html(need):
    """客户侧展示：字段 + 提取结果 + 原文依据 + 置信度；低置信度提示确认/修改。不含任何内部判断。"""
    rows, low = "", []
    for k in ["出发城市", "出行日期", "出行人数", "预算", "目的地", "兴趣", "偏好", "排除条件"]:
        f = need[k]
        v = f["value"]
        v = "、".join(v) if isinstance(v, list) else (f"{v}{need['预算'].get('per', '')}" if k == "预算" and isinstance(v, int) else (f"{v}人" if k == "出行人数" and isinstance(v, int) else v))
        col = {"高": "#16a34a", "中": "#ca8a04", "低": "#dc2626"}[f["conf"]]
        name = "出行时间" if k == "出行日期" else k
        flag = ""
        if f["conf"] == "低" and not (k in ("偏好", "排除条件") or (k == "目的地" and v == "未指定")):
            low.append(name)
            flag = " ⚠请确认/修改"
        rows += f'<tr><td {TD}>{name}</td><td {TD}><b>{esc(str(v))}</b></td><td {TD}>{esc(f["evidence"])}</td><td {TD} style="color:{col}">{f["conf"]}{flag}</td></tr>'
    tip = f'<p style="color:#dc2626;font-size:13px">以下信息置信度较低，请核对并在下方修改：{"、".join(low)}</p>' if low else ""
    return f'<table style="border-collapse:collapse;font-size:13px"><tr><th {TH}>字段</th><th {TH}>提取结果</th><th {TH}>原文依据</th><th {TH}>置信度</th></tr>{rows}</table>{tip}'


# =====================================================================
# 10. 界面（四个页面）
# =====================================================================
EXAMPLES = [
    "我们两个人国庆从上海去首尔，预算6000元/人，喜欢美食和购物，希望住交通方便的酒店，不想坐太早的航班。",
    "国庆两个人从上海出发，预算5000元/人，喜欢美食和自然风景，不知道去哪。",
    "首尔我很想去，但是机票太贵了，我想再看看",
    "我之前想去首尔，但是现在确定不去了。",
    "我准备去不去还没决定，但是不希望收到任何营销推荐，也不要联系我。",
    "10月10日确定去首尔，正在看机票和酒店。",
]


def build_ui():
    import gradio as gr
    refresh_market_state()
    dest_choices = ["未指定（系统推荐）"] + [p["dest"] for p in STATE["pool"]] + sorted(d for d in STATE["features"] if d not in {p["dest"] for p in STATE["pool"]})

    with gr.Blocks(title="航旅智慧营销决策台") as demo:
        gr.Markdown("## 航旅智慧营销决策台（内部侧 · 国内航线为主 · 国际航线扩展）\n客户需求（客户侧提交）＋ 群体层市场信号 → "
                    "市场数据 → 市场洞察 → 需求提取 → 动态画像 → 航线/目的地匹配 → 触达决策 → 渠道与时机 → AIGC内容 → 人工审核 → 效果评估 → 数据反馈 → 下一轮优化")
        st_last = gr.State(None)
        st_prev = gr.State(None)
        with gr.Tabs():
            with gr.Tab("市场洞察"):
                b_upd = gr.Button("更新市场数据", variant="primary")
                m_status = gr.HTML(status_html())
                m_top = gr.HTML(top3_html())
                with gr.Accordion("▼ 查看全部候选目的地（点击某一行查看其关键词明细）", open=False):
                    m_all = gr.Dataframe(final_table(), interactive=False, wrap=True)
                m_kw = gr.HTML(keyword_detail_html(default_dest()))
                st_dest = gr.State(default_dest())
            with gr.Tab("客户需求"):
                gr.Markdown("**个体层 · 客户需求**：这里显示客户在“结伴出游模拟器”中确认并提交的真实结构化需求（非固定测试数据）。")
                with gr.Row():
                    r_pick = gr.Dropdown([], label="客户提交的需求（默认最新）", allow_custom_value=True)  # 选项在运行时动态刷新，需关闭服务端固定选项校验
                    b_rl = gr.Button("刷新客户需求")
                r_html = gr.HTML(request_html(None))
                b_ra = gr.Button("导入并分析 → 旅客画像 / 智能营销", variant="primary")
            with gr.Tab("旅客画像"):
                p_html = gr.HTML(profile_html(None))
            with gr.Tab("智能营销"):
                gr.Markdown("内部手动录入（测试/电话咨询场景）。客户侧提交的需求请在“客户需求”页导入分析。")
                i_text = gr.Textbox(label="自然语言需求输入（可直接粘贴用户原话，系统自动提取并辅助填充下方字段）", lines=3)
                gr.Examples(EXAMPLES, inputs=i_text)
                with gr.Row():
                    i_city = gr.Textbox(label="营销任务：出发城市", value="上海")
                    i_hol = gr.Dropdown(["未指定"] + list(dict.fromkeys(HOLIDAYS)), value="未指定", label="营销任务：假期/出行时间")
                    i_dest = gr.Dropdown(dest_choices, value=dest_choices[0], label="目的地", allow_custom_value=True)
                    i_bud = gr.Number(label="预算（元/人）", value=None)
                    i_n = gr.Number(label="出行人数", value=None, precision=0)
                with gr.Row():
                    i_tags = gr.CheckboxGroup(TAGS, label="兴趣标签")
                    i_date = gr.Textbox(label="计划出行日期", placeholder="例如 2026-10-10")
                    i_intl = gr.Checkbox(label="用户明确选择国际旅行（国际扩展场景）", value=False)
                with gr.Row():
                    b_clear = gr.Button("清空")
                    b_go = gr.Button("提交分析", variant="primary")
                o_html = gr.HTML()
                o_aigc = gr.Textbox(label="10. AIGC营销文案（AIGC生成内容，需人工审核后使用）", lines=9)
                with gr.Row():
                    i_review = gr.Radio(["待审核", "审核通过", "退回修改", "决定不触达"], value="待审核", label="人工审核")
                    i_note = gr.Textbox(label="审核意见（可选）")
                o_review = gr.Markdown("审核状态：待审核")
            with gr.Tab("效果评估"):
                e_html = gr.HTML(effect_html(None, "待审核")[0])
                e_fig = gr.Plot()

        b_upd.click(refresh_market_ui, None, [m_top, m_all, m_kw, st_dest, m_status])

        def pick(df, evt: gr.SelectData):
            row = evt.index[0] if isinstance(evt.index, (list, tuple)) else evt.index
            try:
                dest = df.iloc[row]["目的地"]
            except Exception:
                dest = final_table().iloc[row]["目的地"]
            return keyword_detail_html(dest), dest

        m_all.select(pick, m_all, [m_kw, st_dest])

        def go(text, city, dest, bud, n, tags, date, intl, hol, last):
            R = run_pipeline(text, city, dest, bud, n, tags, date, intl, hol)
            need = R["need"]
            fill = lambda k, cur: need[k]["value"] if need[k]["conf"] == "高" and need[k]["value"] not in ("未指定",) else cur
            nd = need["出行日期"]["date"]
            return (result_html(R), R["content"], profile_html(R, last), R, last, "待审核", "审核状态：**待审核**（AIGC内容需人工审核后方可使用）",
                    fill("出发城市", city), (R["focus"] if R["specified"] else dest),
                    fill("预算", bud), fill("出行人数", n),
                    [x for x in (need["兴趣"]["value"] if isinstance(need["兴趣"]["value"], list) else []) if x in TAGS] or tags,
                    nd.isoformat() if nd else date, *effect_html(R, "待审核"))

        def req_choices():
            rs = load_requests()
            ch = [(f"{r['id']}｜{r['需求生成时间']}｜{r['出发城市']}→{r['目的地']}", r["id"]) for r in reversed(rs)]
            return ch, (ch[0][1] if ch else None)

        def req_refresh():
            ch, v = req_choices()
            rs = {r["id"]: r for r in load_requests()}
            return gr.update(choices=ch, value=v), request_html(rs.get(v))

        def req_pick(v):
            return request_html(next((r for r in load_requests() if r["id"] == v), None))

        def req_analyze(v, last):
            req = next((r for r in load_requests() if r["id"] == v), None)
            if not req:
                gr.Warning("暂无客户需求可分析")
                return (gr.update(),) * 15
            R = run_from_request(req)
            need = R["need"]
            nd = need["出行日期"]["date"]
            return (result_html(R), R["content"], profile_html(R, last), R, last, "待审核", "审核状态：**待审核**（AIGC内容需人工审核后方可使用）",
                    need["出发城市"]["value"], (R["focus"] if R["specified"] else dest_choices[0]),
                    need["预算"]["value"] if isinstance(need["预算"]["value"], (int, float)) else None,
                    need["出行人数"]["value"] if isinstance(need["出行人数"]["value"], int) else None,
                    [x for x in (need["兴趣"]["value"] if isinstance(need["兴趣"]["value"], list) else []) if x in TAGS],
                    nd.isoformat() if nd else "", *effect_html(R, "待审核"))

        go_outs = [o_html, o_aigc, p_html, st_last, st_prev, i_review, o_review, i_city, i_dest, i_bud, i_n, i_tags, i_date, e_html, e_fig]
        b_rl.click(req_refresh, None, [r_pick, r_html])
        r_pick.change(req_pick, r_pick, r_html)
        b_ra.click(req_analyze, [r_pick, st_last], go_outs)
        demo.load(req_refresh, None, [r_pick, r_html])

        b_go.click(go, [i_text, i_city, i_dest, i_bud, i_n, i_tags, i_date, i_intl, i_hol, st_last],
                   [o_html, o_aigc, p_html, st_last, st_prev, i_review, o_review, i_city, i_dest, i_bud, i_n, i_tags, i_date, e_html, e_fig])

        def review(status, note, R):
            if not R:
                return "审核状态：请先生成方案", *effect_html(None, status), profile_html(None), None
            if R["decision"] == "建议不触达":
                return "审核状态：无需审核（建议不触达是正式业务结果，未生成营销内容）", *effect_html(R, status), profile_html(R), R
            msg = {"待审核": "待审核", "审核通过": "审核通过 ✔（可进入模拟触达，并在“效果评估”页查看模拟效果）",
                   "退回修改": "退回修改 ✎（暂不投放，修改后重新审核）", "决定不触达": "决定不触达 ⛔（正式业务结果，不投放）"}[status]
            R2 = dict(R)
            E = simulate_effect(R, status)
            R2["feedback"] = feedback_items(E) if E and not E["skip"] else []
            return (f"审核状态：**{msg}**" + (f"｜意见：{note}" if note else ""), *effect_html(R, status), profile_html(R2, None), R2)

        i_review.change(review, [i_review, i_note, st_last], [o_review, e_html, e_fig, p_html, st_last])
        b_clear.click(lambda: ("", "上海", dest_choices[0], None, None, [], "", False, "", ""), None,
                      [i_text, i_city, i_dest, i_bud, i_n, i_tags, i_date, i_intl, o_html, o_aigc])
    return demo


CUSTOMER_WORDS = ["百度指数", "市场热度", "粗筛", "精筛", "营销渠道", "触达", "AIGC", "人工审核", "ROI", "效果评估", "决策台", "评分"]  # 客户侧页面禁止出现的内部词（测试用）


def build_customer_ui():
    """客户侧：结伴出游模拟器。只采集需求、提取、确认、提交；不展示任何内部营销数据或判断。"""
    import gradio as gr
    with gr.Blocks(title="结伴出游模拟器") as demo:
        gr.Markdown("## 结伴出游模拟器\n告诉我们您的出游想法，可以直接用一句话描述，也可以逐项填写。提交后我们会根据您的需求提供服务。")
        st_ext = gr.State(None)
        c_text = gr.Textbox(label="说说您的出游需求（自然语言）", lines=3, placeholder="例如：我们两个人国庆从上海出发，预算5000元/人，喜欢美食和自然风景，还没有确定去哪，不想坐太早的航班。")
        with gr.Row():
            c_city = gr.Textbox(label="出发城市")
            c_time = gr.Textbox(label="出行时间/假期", placeholder="例如 国庆 或 2026-10-10")
            c_n = gr.Number(label="出行人数", value=None, precision=0)
            c_bud = gr.Number(label="预算（元/人）", value=None)
        with gr.Row():
            c_dest = gr.Textbox(label="目的地（没想好请留空或填“未指定”）", value="未指定")
            c_tags = gr.CheckboxGroup(TAGS, label="兴趣")
        with gr.Row():
            c_pref = gr.Textbox(label="偏好（如 交通方便）")
            c_excl = gr.Textbox(label="排除条件（如 不要太早的航班）")
        b_ext = gr.Button("提取我的需求", variant="primary")
        c_tbl = gr.HTML()
        c_ok = gr.Checkbox(label="我已核对以上信息，确认无误", value=False)
        b_sub = gr.Button("确认并提交")
        c_msg = gr.Markdown()

        def do_extract(text, city, time_, n, bud, dest, tags, prefs, excl):
            need = customer_extract(text, city, time_, n, bud, dest, tags, prefs, excl)
            f = need_to_fields(need)
            return (customer_table_html(need), f["city"], f["time"], f["n"], f["bud"], f["dest"], f["tags"], f["prefs"], f["excl"], need, False)

        b_ext.click(do_extract, [c_text, c_city, c_time, c_n, c_bud, c_dest, c_tags, c_pref, c_excl],
                    [c_tbl, c_city, c_time, c_n, c_bud, c_dest, c_tags, c_pref, c_excl, st_ext, c_ok])

        def do_submit(text, city, time_, n, bud, dest, tags, prefs, excl, ok, need):
            if not ok:
                return "请先点击“提取我的需求”核对信息，并勾选“我已核对以上信息，确认无误”。"
            if not (str(text or "").strip() or city or time_ or n or bud or tags):
                return "请至少填写一项出游需求。"
            req = build_request(text, city, time_, n, bud, dest, tags, prefs, excl, need)
            rid = save_request(req)
            return f"✅ 您的出游需求已生成并提交（编号：{rid}）。感谢您的信任！"

        b_sub.click(do_submit, [c_text, c_city, c_time, c_n, c_bud, c_dest, c_tags, c_pref, c_excl, c_ok, st_ext], c_msg)
    return demo


LANDING = """<!doctype html><html lang="zh"><head><meta charset="utf-8"><title>航旅智慧营销 · 双端入口</title>
<style>body{font-family:system-ui,"PingFang SC","Microsoft YaHei";background:#f9fafb;margin:0;padding:48px}h1{font-size:24px}
.row{display:flex;gap:24px;flex-wrap:wrap}a.card{display:block;width:320px;padding:24px;border-radius:14px;background:#fff;border:1px solid #e5e7eb;color:#111;text-decoration:none}
a.card:hover{border-color:#f97316}small{color:#6b7280}</style></head><body><h1>航旅智慧营销 Demo · 请选择入口</h1><div class="row">
<a class="card" href="/customer/"><h2>客户入口</h2><b>结伴出游模拟器</b><p><small>客户表达出游需求、确认并提交。</small></p></a>
<a class="card" href="/internal/"><h2>内部入口</h2><b>航旅智慧营销决策台</b><p><small>营销人员使用：市场洞察、客户需求、旅客画像、智能营销、效果评估。</small></p></a>
</div></body></html>"""


def build_site():
    """同一服务的两个入口：/customer（客户侧）与 /internal（内部侧），通过共享的需求存储真实传递数据。"""
    import gradio as gr
    from fastapi import FastAPI
    from fastapi.responses import HTMLResponse
    api = FastAPI()

    @api.get("/", response_class=HTMLResponse)
    def landing():
        return LANDING

    gr.mount_gradio_app(api, build_customer_ui(), path="/customer")
    gr.mount_gradio_app(api, build_ui(), path="/internal")
    return api


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "7860"))
    print(f"客户入口：http://127.0.0.1:{port}/customer/    内部入口：http://127.0.0.1:{port}/internal/    入口选择页：http://127.0.0.1:{port}/")
    uvicorn.run(build_site(), host="127.0.0.1", port=port)
