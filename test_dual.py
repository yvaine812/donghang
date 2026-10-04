# -*- coding: utf-8 -*-
"""双端测试 T01~T07：客户侧提取→确认→真实传入内部决策台。运行：python test_dual.py"""
import os, tempfile, json, time
import app

app.REQ_FILE = os.path.join(tempfile.mkdtemp(), "customer_requests.json")  # 测试用临时文件，不污染项目数据
app.refresh_market_state()
assert app.STATE["ok"], app.STATE["error"]
results = []


def check(tid, desc, ok, extra=""):
    results.append(bool(ok))
    print(f"{'PASS' if ok else 'FAIL'} {tid} {desc} {extra}")


def customer_submit(text, **kw):
    """模拟客户侧完整流程：提取 → （客户核对，可修改）→ 确认提交；返回内部侧读取到的请求。"""
    need = app.customer_extract(text)
    f = app.need_to_fields(need)
    f.update(kw)  # 客户在确认页修改字段
    req = app.build_request(text, f["city"], f["time"], f["n"], f["bud"], f["dest"], f["tags"], f["prefs"], f["excl"], need)
    rid = app.save_request(req)
    got = next(r for r in app.load_requests() if r["id"] == rid)  # 内部侧从共享存储读取
    return got, app.run_from_request(got), need


# T01 目的地未指定：客户侧 → 内部侧
t1 = "国庆两个人从上海出发，预算5000元/人，喜欢美食和自然风景，还不知道去哪。"
req, R, need = customer_submit(t1)
rt = [r["dest"] for r in R["task"]["routes"]]
pool = {p["dest"] for p in app.STATE["pool"]}
need_keys = {"出发城市", "出行时间", "出行人数", "预算", "目的地", "兴趣", "偏好", "排除条件", "原始自然语言", "字段置信度", "需求生成时间"}
check("T01", "目的地未指定：提取→确认→传入决策台→3~5条国内候选航线", need_keys <= set(req) and req["目的地"] == "未指定" and req["原始自然语言"] == t1
      and R["source"]["id"] == req["id"] and not R["specified"] and 3 <= len(rt) <= 5 and "首尔" not in rt
      and all(d in pool for d in rt) and all(app.dest_type(d) == "国内" for d in rt) and all(r["reasons"] for r in R["task"]["routes"])
      and R["profile"]["name"] and "客户侧" in R["source"]["kind"], f"｜航线 {['上海→'+d for d in rt]}｜画像 {R['profile']['name']}")
# 客户修改低置信度字段后，传到内部的是修改后的值
req2, R2, _ = customer_submit(t1, bud=3000)
check("T01b", "客户修改预算后，内部收到的是客户确认值", req2["预算"] == 3000 and R2["need"]["预算"]["value"] == 3000)

# T02 指定目的地
req, R, _ = customer_submit("上海去首尔，两个人，预算6000元，喜欢美食购物。")
check("T02", "指定目的地：直接分析上海→首尔", R["specified"] and R["focus"] == "首尔" and [r["dest"] for r in R["task"]["routes"]] == ["首尔"] and req["出发城市"] == "上海")

# T03~T05 触达决策
req, R, _ = customer_submit("首尔我很想去，但是机票太贵了，我想再看看。")
check("T03", "价格犹豫 → 谨慎触达", R["decision"] == "谨慎触达", f"｜{R['decision']}")
req, R, _ = customer_submit("我现在不去了，也不要给我任何营销推荐。")
check("T04", "明确拒绝 → 建议不触达", R["decision"] == "建议不触达", f"｜{R['decision']}")
req, R, _ = customer_submit("10月10日确定去首尔，正在看机票和酒店。")
check("T05", "明确购买 → 建议触达", R["decision"] == "建议触达", f"｜{R['decision']}")

# T06 动态变化（走客户侧→内部侧全流程）
base = "国庆两个人从上海出发，喜欢美食和购物，预算{}元/人。"
_, Ra, _ = customer_submit(base.format(6000))
_, Rb, _ = customer_submit(base.format(3000))
ra, rb = [r["dest"] for r in Ra["task"]["routes"]], [r["dest"] for r in Rb["task"]["routes"]]
check("T06", "预算6000→3000：画像、候选航线、营销策略合理变化", Ra["profile"]["name"] != Rb["profile"]["name"] and ra != rb
      and (Ra["strategy"] != Rb["strategy"] or Ra["content"] != Rb["content"]), f"｜{ra} → {rb}｜{Ra['profile']['name']} → {Rb['profile']['name']}")

# T07 更新市场数据
s0 = (app.STATE["time"], app.STATE["version"], id(app.STATE["pool"]))
time.sleep(1.1)
out = app.refresh_market_ui()
s1 = (app.STATE["time"], app.STATE["version"], id(app.STATE["pool"]))
st = out[4]
check("T07", "更新市场数据：时间/次数变化，最终候选池被刷新，显示状态", s1[0] != s0[0] and s1[1] == s0[1] + 1 and s1[2] != s0[2] and 8 <= len(app.STATE["pool"]) <= 12
      and "已加载最新数据" in st and app.STATE["time"] in st, f"｜{s0[0]}→{s1[0]}｜次数 {s0[1]}→{s1[1]}｜候选池 {len(app.STATE['pool'])} 个")

# 补充：客户侧不泄露内部信息；热词表不在请求时读取；两轮筛选
cust = app.build_customer_ui()
cfg = json.dumps(cust.get_config_file(), ensure_ascii=False, default=str)
leak = [w for w in app.CUSTOMER_WORDS if w in cfg]
check("T08", "客户侧页面不含任何内部营销信息", not leak, f"｜命中:{leak}")
calls = []
orig = app.load_hotwords
app.load_hotwords = lambda: calls.append(1) or orig()
customer_submit(t1)
app.load_hotwords = orig
check("T09", "每个客户请求不重新读取原始热词统计表", not calls)
src = open(app.__file__, encoding="utf-8").read()
check("T10", "只有两轮筛选，无第三轮；最终候选池8~12个", "第三轮" not in src.replace("不设第三轮", "").replace("没有第三轮", "").replace("不增加第三轮", "") and 8 <= len(app.STATE["pool"]) <= 12)
print(f"\n{sum(results)}/{len(results)} 通过")
assert all(results)
