#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
lesson-pack 档A 生成器：把单源内容 JSON（结构见 templates/lesson-content.schema.md）
渲染成一份「自包含 HTML 教学包」——纸质质感、浏览器直接看 / 转 PDF / 打印。
视觉严格贴合 SKILL.md「视觉与版式」铁律：米灰底 + 暖白纸面 + 楷宋衬线 +
主体黑墨 + 红笔(#b3261e)标时间/重难点 + 蓝(#1a4f8b)标栏目 + 卷子样式 + @media print 分页。
与档B(lesson_pack_gen.py)同源，内容一致。

用法：
  python lesson_pack_html.py --data 内容.json [--out 教学包.html]
"""
import os, sys, argparse, json, html, datetime

def esc(s):
    return html.escape(str(s), quote=False)

def check_time(L):
    """时间账硬校验（quality-checklist A 项）：时间总表求和 == 课时，且环节分求和也 == 课时。
    档A 此前完全没有校验，时间账错了也静默产出——与 quality-checklist 的"A 项不过必须返工"矛盾。
    与 lesson_pack_gen.py 的 check_time 保持同一套规则，两档口径一致。"""
    issues = []
    total = None
    try:
        total = sum(int(m) for _, m in L.get('时间总表') or [])
    except Exception as e:
        issues.append(f"时间总表格式异常（应为 [[环节名, 分钟], …]）：{e}")
    if total is not None:
        try:
            if total != int(L.get('课时')):
                issues.append(f"时间总表合计 {total}′ ≠ 课时 {L.get('课时')}′")
        except Exception:
            issues.append(f"课时字段缺失或非数字：{L.get('课时')!r}")
    try:
        ph_sum = sum(int(p['分']) for p in L.get('环节') or [])
        if ph_sum != int(L.get('课时')):
            issues.append(f"各环节分钟合计 {ph_sum}′ ≠ 课时 {L.get('课时')}′")
    except Exception as e:
        issues.append(f"环节分钟（环节[].分）格式异常：{e}")
    return (len(issues) == 0), issues


def exercise_minutes(L):
    """课堂练习卷建议用时，取值优先级：lesson.练习时长 > 含"练习/巩固/训练/检测"的环节分钟
    > 课时/3（夹在 5–20′）。旧实现直接写整节课时长（如 40′），明显不对。
    与 lesson_pack_gen.py 的 exercise_minutes 规则一致，保证两档显示同一个数。"""
    v = L.get('练习时长')
    if v:
        try:
            return int(v)
        except Exception:
            pass
    for ph in (L.get('环节') or []):
        if any(k in str(ph.get('名', '')) for k in ('练习', '巩固', '训练', '检测')):
            try:
                return int(ph['分'])
            except Exception:
                pass
    try:
        return min(20, max(5, int(L.get('课时', 40)) // 3))
    except Exception:
        return 15


def fmt_time_ts(ts):
    """时间总表 → '导入 4′ ｜ 新授 18′…'"""
    return " ｜ ".join(f"{n} <b class='red'>{m}</b>′" for n, m in ts)

CSS = """
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:#e8e5de;color:#2a2826;
    font-family:"Kaiti SC","STKaiti","KaiTi",Georgia,"Songti SC","SimSun",serif;
    line-height:1.85;padding:0 0 60px;-webkit-font-smoothing:antialiased}
  .deck{max-width:820px;margin:0 auto;padding:0 14px}
  .sheet{background:#fdfdf9;border:1px solid #d8d4c9;
    box-shadow:0 1px 4px rgba(90,80,60,.10);padding:56px 62px 52px;margin:26px 0;position:relative}
  @media(max-width:600px){.sheet{padding:34px 22px 30px}}
  .sheet::before{content:"";position:absolute;left:30px;top:0;bottom:0;border-left:1px dashed #d9d4c6}
  h1,h2,h3{font-weight:700;color:#1f1d1a;line-height:1.4}
  h1{font-size:24px;text-align:center;letter-spacing:2px}
  h2{font-size:20px;margin:28px 0 4px;padding-bottom:5px;border-bottom:1.5px solid #333}
  h3{font-size:16px;margin:20px 0 6px}
  p,li,td,th,div{font-size:15px}
  .center{text-align:center}
  .subtitle{text-align:center;color:#555;font-size:15px;margin:4px 0 2px}
  .meta-line{text-align:center;color:#777;font-size:13px;letter-spacing:2px;margin-bottom:10px}
  .red{color:#b3261e;font-weight:700}
  .blue{color:#1a4f8b;font-weight:700}
  table{width:100%;border-collapse:collapse;margin:8px 0;font-size:14.5px}
  th,td{border:1px solid #6b6b6b;padding:8px 10px;vertical-align:top;text-align:left}
  th{background:#f3f0e6;font-weight:700}
  td.rowlabel{background:#faf7ec;font-weight:700;width:92px;text-align:center;vertical-align:middle}
  ol.flow,ul.plain{margin:6px 0 6px 22px}
  ol.flow li,ul.plain li{margin:5px 0}
  .step{font-weight:700;font-size:16px;margin:22px 0 6px;color:#1f1d1a;
    border-left:4px solid #b3261e;padding-left:10px;clear:both;overflow:hidden}
  .step .t{float:right;color:#b3261e;font-weight:700;border:1px solid #b3261e;border-radius:3px;padding:0 8px;font-size:13px}
  .board{border:1.5px solid #333;padding:14px 18px;font-size:15px;line-height:1.9;margin:10px 0;background:#fffdf7}
  .board b{display:block;border-bottom:1px solid #999;padding-bottom:2px;margin-bottom:6px}
  .time-sum{background:#f6f2e6;border:1px solid #cfc9b6;padding:10px 14px;font-size:14.5px;margin:12px 0}
  .time-sum b{color:#b3261e}
  .paper h2{text-align:center;border:none;letter-spacing:3px}
  .paper .hd{display:flex;justify-content:space-between;font-size:14px;color:#333;margin:10px 0 18px}
  .q{display:flex;gap:8px;margin:13px 0}
  .q .no{font-weight:700;min-width:18px}
  .q .body{flex:1}
  .opt{display:flex;flex-wrap:wrap;gap:22px;margin-top:6px}
  .answer-line{height:28px;border-bottom:1px solid #bbb;margin:4px 0}
  .fill{display:inline-block;border-bottom:1px solid #666;min-width:90px;height:1.05em;vertical-align:bottom}
  .fillline{display:inline-block;border-bottom:1px solid #666;height:1.1em;vertical-align:bottom}
  hr.dashed{border:none;border-top:1.5px dashed #aaa;margin:22px 0}
  .page-note{color:#8a8a80;font-size:12.5px;text-align:center;letter-spacing:1px;margin-top:26px}
  @media print{body{background:#fff}.sheet{box-shadow:none;border:none;page-break-after:always;margin:0}.sheet::before{display:none}}
"""

def sheet_design(L):
    """纸1 教学设计"""
    b = L.get("板书") or {}   # 板书缺失时降级为只显示课题，不再 KeyError 崩溃
    objs = []
    objs.append('<section class="sheet">')
    objs.append(f"<h1>教 学 设 计</h1>")
    objs.append(f'<div class="meta-line">{esc(L["教材"])} · {esc(L["章节"])}</div>')
    objs.append(f'<div class="center" style="font-size:17px;font-weight:700">课题：{esc(L["课题"])}（{esc(L["课时标签"])} · {esc(L["课型"])}）</div>')
    objs.append('<table style="margin-top:16px">')
    objs.append(f'<tr><td class="rowlabel">学科 / 年级</td><td>{esc(L["学科"])} · {esc(L["年级"])}</td>'
                f'<td class="rowlabel">课型</td><td>{esc(L["课型"])}</td></tr>')
    objs.append(f'<tr><td class="rowlabel">教材位置</td><td>{esc(L["教材"])} {esc(L["章节"])}</td>'
                f'<td class="rowlabel">课时</td><td><span class="red">{L["课时"]} 分钟</span></td></tr>')
    objs.append('</table>')
    # 目标
    objs.append("<h2>一、教学目标</h2>")
    for cat, items in L["教学目标"]:
        objs.append(f"<p><span class='blue'>{esc(cat)}：</span></p>")
        objs.append("<ol class='flow'>" + "".join(f"<li>{esc(i)}</li>" for i in items) + "</ol>")
    # 重难点
    objs.append("<h2>二、教学重难点</h2>")
    objs.append(f"<p><b>重点：</b>{esc(L['重难点'][0])}</p>")
    objs.append(f"<p><b>难点：</b>{esc(L['重难点'][1])}</p>")
    objs.append(f"<p><b class='red'>突破：</b>{esc(L['重难点'][2])}</p>")
    # 学情
    objs.append("<h2>三、学情分析</h2>")
    objs.append(f"<p>{esc(L['学情'])}</p>")
    # 环节
    objs.append("<h2>四、教学环节</h2>")
    total = sum(m for _, m in L["时间总表"])
    objs.append(f'<div class="time-sum">时间分配：{fmt_time_ts(L["时间总表"])} —— 合计 <b>{total}</b>′</div>')
    for i, ph in enumerate(L["环节"]):
        objs.append(f'<div class="step">{esc(ph["名"])} <span class="t">{ph["分"]}′</span></div>')
        objs.append("<table>")
        for k, v in ph["rows"]:
            objs.append(f"<tr><td class='rowlabel'>{esc(k)}</td><td>{esc(v).replace(chr(10), '<br>')}</td></tr>")
        objs.append("</table>")
    # 板书
    objs.append("<h2>五、板书设计</h2>")
    objs.append('<div class="board">')
    objs.append(f"<b>{esc(b.get('title') or L['课题'])}</b>")
    if b.get("concept"):
        objs.append(f'<div style="margin:2px 0">{esc(b["concept"])}</div>')
    objs.append('<div style="display:flex;gap:24px;margin-top:8px">')
    cols = []
    for col in ("left", "mid", "right"):
        cols.append(f"<div style='flex:1'>" + "<br>".join(esc(x) for x in (b.get(col) or [])) + "</div>")
    objs.append("".join(cols))
    objs.append("</div></div>")
    objs.append('<div class="page-note">—— 教学设计 · 第 1 页 ——</div>')
    objs.append("</section>")
    return "\n".join(objs)

def hd_line(extra_left="", extra_right=""):
    """班级/姓名/日期 + 可选右信息"""
    r = ("<div class='hd'><span>班级<span class='fill'></span>　姓名<span class='fill'></span>"
         f"　日期<span class='fill'></span>{extra_left}</span><span>{extra_right}</span></div>")
    return r

def q_block(no, body_html, lines=1):
    """一道题：题号 + 题干 + 作答留空线"""
    return (f"<div class='q'><div class='no'>{no}.</div><div class='body'>{body_html}"
            + "".join('<div class="answer-line"></div>' for _ in range(lines)) + "</div></div>")

def body_multi(text):
    """题面多行处理：____ / ＿ 视为填空线，其余按行转 <br>。
    顺序关键：先 esc 转义题干里的 <>&，再做填空线替换——此时插入的是安全 HTML，不可再转义。
    若反序会 esc 把刚插的 <span> 也转义，导致页面漏出原始 html 标签文本。"""
    import re
    safe = esc(text).replace(chr(10), "<br>")
    safe = re.sub(r"＿+", lambda m: f"<span class='fill' style='min-width:{max(24, len(m.group())*18)}px'></span>", safe)
    safe = re.sub(r"_{2,}", lambda m: f"<span class='fill' style='min-width:{max(24, len(m.group())*9)}px'></span>", safe)
    return safe

def sheet_worksheet(L, tasks):
    """纸2 学习任务单(学案)"""
    o = ['<section class="sheet paper">', "<h2>学习任务单</h2>",
         f'<div class="subtitle">{esc(L["课题"])} · 课堂用</div>',
         hd_line(extra_right="座号<span class='fill' style='min-width:40px'></span>")]
    goals = L["教学目标"]
    # 取法与档B(make_worksheet_docx)一致：每类最多前 2 条、总共最多 4 条。
    # 旧实现只取"前 2 类各 1 条"，导致两档学案目标条数不一致。
    _sel = []
    for _cat, _its in goals:
        for _g in _its[:2]:
            if len(_sel) < 4:
                _sel.append(_g)
    goal_sum = "　".join("□ " + i for i in _sel)
    o.append(f"<p><b>学习目标：</b>学完我能——{goal_sum}</p>")
    for hname, no, body in tasks:
        o.append(f"<h3>{esc(hname)}</h3>")
        # 首行为题干，换行后用作答留空
        o.append(q_block(no, body_multi(body), lines=0))
        # 题干里已带 ____ 填空线时不再另加整行作答空，避免重复留白
        if not ("＿" in body or "__" in body):
            o.append('<div class="answer-line"></div>')
    o.append('<div class="page-note">—— 学习任务单 · 发给学生当堂用 ——</div></section>')
    return "\n".join(o)

def sheet_exercise(L, exercises):
    """纸3 课堂分层练习(含答案)"""
    o = ['<section class="sheet paper">', "<h2>课堂分层练习</h2>",
         f'<div class="subtitle">{esc(L["课题"])} · 附答案（老师用，答案不下发）</div>',
         hd_line(extra_right=f"时间 {exercise_minutes(L)}′")]
    ans_map = []
    for group, items in exercises:
        o.append(f'<p style="color:#777">{esc(group)}</p>')
        for no, stem, ans in items:
            o.append(q_block(no, body_multi(stem), lines=2))
            ans_map.append(f"<p>{esc(no)}. {esc(ans)}</p>")
    o.append('<hr class="dashed">')
    o.append('<h3 style="color:#b3261e">参考答案（老师保留，勿随卷下发）</h3>')
    o.append("".join(ans_map))
    o.append('<div class="page-note">—— 课堂分层练习 · 附答案 ——</div></section>')
    return "\n".join(o)

def sheet_homework(L, homework):
    """纸4 课后作业(含答案)"""
    o = ['<section class="sheet paper">', "<h2>课后作业</h2>",
         f'<div class="subtitle">{esc(L["课题"])} · {esc(L["作业说明"])}</div>',
         hd_line(extra_right="上交时间：<span class='fill' style='min-width:80px'></span>")]
    ans_map = []
    for group, items in homework.items():
        o.append(f'<p style="color:#777">{esc(group)}</p>')
        for no, stem, ans in items:
            o.append(q_block(no, body_multi(stem), lines=1))
            ans_map.append(f"<p>{esc(no)}. {esc(ans)}</p>")
    o.append('<hr class="dashed">')
    o.append('<h3 style="color:#b3261e">参考答案（老师用）</h3>')
    o.append("".join(ans_map))
    o.append('<div class="page-note">—— 课后作业 ——</div></section>')
    return "\n".join(o)

# lesson 必填字段（结构见 templates/lesson-content.schema.md）。
# 缺字段集中报错，避免渲染到一半 KeyError 崩溃、用户拿到半截 HTML。
BASE_REQUIRED = ['学科', '年级', '教材', '章节', '课题', '课型', '课时',
                 '课时标签', '教学目标', '重难点', '学情', '时间总表', '环节']


def render(data):
    L = data.get("lesson")
    if not L:
        raise SystemExit("✗ 内容 JSON 缺顶层 'lesson' 键（结构见 templates/lesson-content.schema.md）")
    miss = [k for k in BASE_REQUIRED
            if k not in L or (isinstance(L[k], (str, list, dict)) and len(L[k]) == 0)]
    if miss:
        raise SystemExit("✗ 内容 JSON 的 lesson 缺少必填字段：" + "、".join(miss) +
                         "（结构见 templates/lesson-content.schema.md）")
    ok, issues = check_time(L)
    if not ok:
        print("  ⚠ 时间账不平（quality-checklist A 项，请修订内容 JSON）：")
        for it in issues:
            print("      - " + it)
    if data.get("homework") and not L.get("作业说明"):
        raise SystemExit("✗ 提供了 homework，但 lesson 缺 '作业说明'（作业量说明）")
    title = f'{L["课题"]} · 教案与配套（{L["教材"]} · {L["课时标签"]} · {L["课时"]}分钟）'
    parts = [sheet_design(L)]
    if data.get("worksheet_tasks"):
        parts.append(sheet_worksheet(L, data["worksheet_tasks"]))
    if data.get("exercises"):
        parts.append(sheet_exercise(L, data["exercises"]))
    if data.get("homework"):
        parts.append(sheet_homework(L, data["homework"]))
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    footer = (
        '<div style="text-align:center;color:#777;font-size:13px;padding:6px 14px 0;line-height:1.8">'
        f'本套内容由 lesson-pack 生成，遵循"中国课标适配 · 时长驱动({L["课时"]}′ · 单源派生一致 · 纯本地"原则，未采集任何学生个人信息。<br>'
        "可打印后直接用于教学。想调环节时长 / 换课型 / 改难度，告诉我一句即可局部更新。"
        f'<br>生成于 {esc(now)}</div>')
    return (f'<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1.0">'
            f"<title>{esc(title)}</title><style>{CSS}</style></head>"
            f'<body><div class="deck">' + "\n".join(parts) + footer + "</div></body></html>")

def main():
    ap = argparse.ArgumentParser(description="档A：单源JSON → 自包含HTML教学包")
    ap.add_argument("--data", required=True, help="单源内容JSON(见 templates/lesson-content.schema.md)")
    ap.add_argument("--out", default="教学包.html", help="输出HTML路径")
    ap.add_argument("--strict-time", action="store_true",
                    help="时间账（环节分钟合计≠课时）不平时直接终止，不生成 HTML。"
                         "默认仅告警后继续（与档B 行为一致）。")
    a = ap.parse_args()
    # --out 若指向不存在的目录（如 教学包/咏雪/教学包.html），自动建目录；
    # 档B 有 os.makedirs，档A 没有，同样的路径档B 能出、档A 直接 FileNotFoundError。
    _d = os.path.dirname(os.path.abspath(a.out))
    if _d:
        os.makedirs(_d, exist_ok=True)
    data = json.load(open(a.data, encoding="utf-8"))
    if a.strict_time:
        L0 = data.get("lesson") or {}
        ok, _issues = check_time(L0)
        if not ok:
            raise SystemExit("✗ 已启用 --strict-time，时间账不平即终止，未生成文件。")
    h = render(data)
    with open(a.out, "w", encoding="utf-8") as f:
        f.write(h)
    print("已生成 " + a.out + f"（{len(h)//1024}KB）")
    # 哪些纸没出要说清楚，避免老师以为"整套齐了"结果发现少了作业页
    for key, name in (('worksheet_tasks', '学习任务单'), ('exercises', '课堂分层练习'),
                      ('homework', '课后作业')):
        if not data.get(key):
            print(f"  · 内容 JSON 未提供 '{key}'，已跳过「{name}」页")

if __name__ == "__main__":
    main()
