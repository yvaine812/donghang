# -*- coding: utf-8 -*-
"""T01~T11 场景测试：python test_scenarios.py"""
import app

app.refresh_market_state()
assert app.STATE["ok"], app.STATE["error"]
F = dict(city="上海", dest="未指定（系统推荐）", budget=None, n=None, tags=[], date="", intl=False)


def run(text="", holiday="", **kw):
    a = {**F, **kw}
    return app.run_pipeline(text, a["city"], a["dest"], a["budget"], a["n"], a["tags"], a["date"], a["intl"], holiday)


results = []


def check(tid, desc, R, cond, extra=""):
    ok = bool(cond)
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'} {tid} {desc} → 决策:{R['decision']}｜阶段:{R['stage']}｜目的地:{R['focus']}｜画像:{R['profile']['name']} {extra}")


R = run("准备10月出行，两个人从上海去首尔，预算6000元/人，喜欢美食、购物和娱乐，准备订票。", tags=["美食", "购物", "娱乐"])
check("T01", "上海→首尔 高意向", R, R["decision"] == "建议触达" and R["focus"] == "首尔")
R = run("首尔我很想去，但是机票太贵了，我想再看看")
check("T02", "想去但贵", R, R["decision"] == "谨慎触达" and R["focus"] == "首尔")
R = run("我之前想去首尔，但是现在确定不去了。")
check("T03", "确定不去", R, R["decision"] == "建议不触达")
R = run("我准备去不去还没决定，但是不希望收到任何营销推荐，也不要联系我。")
check("T04", "拒绝营销", R, R["decision"] == "建议不触达" and "第一优先级" in R["reason"])
R = run("最近没有去首尔的计划，暂时不考虑出行。")
check("T05", "暂无计划", R, R["decision"] == "建议不触达")
R = run("10月10日确定去首尔，正在看机票和酒店。")
check("T06", "确定去+看机票", R, R["decision"] == "建议触达" and R["stage"] == "已明确出行/购买阶段")
R = run("想出去旅游，两个人，预算1500元/人，喜欢美食。")
bad = [r for r in R["recs"] if r["budget"] == "不足"]
check("T07", "预算1500", R, R["decision"] == "谨慎触达" and not bad and all(r["budget"] != "不足" for r in R["recs"]),
      f"｜可推荐:{[r['dest']+'('+r['budget']+')' for r in R['recs']]}｜因预算排除:{len(R['excluded'])}个")
t = "国庆两个人从上海出发，预算{}元/人，喜欢美食和购物。"
R1, R2 = run(t.format(6000)), run(t.format(3000))
check("T08", "预算6000→3000 动态变化", R2, R1["profile"]["name"] != R2["profile"]["name"] and
      ([r["dest"] for r in R1["recs"][:3]] != [r["dest"] for r in R2["recs"][:3]] or R1["content"] != R2["content"]),
      f"｜V1:{R1['profile']['name']} 推荐{[r['dest'] for r in R1['recs'][:3]]}｜V2推荐{[r['dest'] for r in R2['recs'][:3]]}")
R = run("国庆两个人从上海出发，预算5000元/人，喜欢美食和自然风景，不知道去哪。")
check("T09", "未指定目的地", R, not R["specified"] and R["focus"] != "首尔" and app.dest_type(R["focus"]) == "国内" and R["recs"][0]["reasons"],
      f"｜推荐:{[r['dest'] for r in R['recs'][:3]]}")
R = run("", holiday="国庆")
rt = [r["dest"] for r in R["task"]["routes"]]
check("T10", "营销任务：上海+国庆 → 3~5条候选航线", R, R["task"]["city"] == "上海" and R["task"]["holiday"] == "国庆" and 3 <= len(rt) <= 5
      and "首尔" not in rt and all(app.dest_type(d) == "国内" for d in rt), f"｜航线:{['上海→'+d for d in rt]}")
R = run("两个人，预算1500元/人，喜欢美食。", holiday="国庆")
pool_names = {p["dest"] for p in app.STATE["pool"]}
check("T11", "候选不足(<3)：不补齐、显示实际数量、给出提示", R, len(R["task"]["routes"]) < 3 and R["task"]["insufficient"]
      and "候选不足" in R["task"]["note"] and all(r["dest"] in pool_names for r in R["task"]["routes"]),
      f"｜实际{len(R['task']['routes'])}条：{R['task']['note']}")
R = run("我们两个人国庆从上海去首尔，预算6000元/人。")
check("T12", "指定目的地：直接分析上海→首尔，不被替换", R, R["specified"] and R["focus"] == "首尔" and [r["dest"] for r in R["task"]["routes"]] == ["首尔"])
R = run("国庆两个人从上海出发，预算5000元/人，喜欢美食和自然风景，不知道去哪。")
check("T13", "无目的地：不默认首尔，目的地来自最终候选池", R, not R["specified"] and R["focus"] in {p["dest"] for p in app.STATE["pool"]} and R["focus"] != "首尔")
R = run("国庆两个人从上海出发，预算5000元/人，喜欢美食和自然风景，还不知道去哪。")
rt = [r["dest"] for r in R["task"]["routes"]]
check("T09b", "未指定目的地(命令原句)：3~5条国内候选航线，每条有理由，不默认首尔", R, R["need"]["目的地"]["value"] == "未指定" and 3 <= len(rt) <= 5
      and "首尔" not in rt and all(app.dest_type(d) == "国内" for d in rt) and all(r["reasons"] for r in R["task"]["routes"]) and R["stage"] == "探索阶段",
      f"｜航线:{['上海→'+d for d in rt]}")
check("T15", "只使用最新版A表(1).md，不读旧版xlsx", R, "A成员数据搜集表(1).md" in app.STATE["source"] and "xlsx" not in app.STATE["source"])
check("T16", "AIGC建议卡含客群/航线/理由/渠道/海报/AIGC标识/审核状态", R, all(k in R["content"] for k in ("【客群】", "【推荐航线】", "【推荐理由】", "【营销渠道】", "【视觉/海报建议】", "AIGC生成，需人工审核", "待审核")))
R1, R2 = run("国庆两个人从上海出发，预算6000元/人，喜欢美食和购物。"), run("国庆两个人从上海出发，预算3000元/人，喜欢美食和购物。")
check("T08b", "预算6000→3000：航线排序/画像/营销策略均变化", R2, [r["dest"] for r in R1["task"]["routes"]] != [r["dest"] for r in R2["task"]["routes"]]
      and R1["profile"]["name"] != R2["profile"]["name"] and (R1["strategy"] != R2["strategy"] or R1["content"] != R2["content"]),
      f"｜6000:{[r['dest'] for r in R1['task']['routes']]}｜3000:{[r['dest'] for r in R2['task']['routes']]}")
# 硬编码检查：app.py 中“首尔/丽江/2980/8500”只允许出现在 EXAMPLES 示例文本里
src = open(app.__file__, encoding="utf-8").read().split("\n")
ex_start = next(i for i, l in enumerate(src) if l.startswith("EXAMPLES"))
ex_end = next(i for i in range(ex_start, len(src)) if src[i].startswith("]"))
bad = [i + 1 for i, l in enumerate(src) if any(k in l for k in ("首尔", "丽江", "2980", "8500")) and not (ex_start <= i <= ex_end)]
results.append(not bad)
print(f"{'PASS' if not bad else 'FAIL'} T14 无首尔/丽江/2980/8500硬编码（仅示例文本允许）{bad}")
# 需求提取示例
N = app.analyze_user_need("我们两个人国庆从上海去首尔，预算6000元/人，喜欢美食和购物，希望住交通方便的酒店，不想坐太早的航班。")
print({k: (v["value"], v["conf"]) for k, v in N.items() if not k.startswith("_")})
print(f"\n{sum(results)}/{len(results)} 通过")
assert all(results)
