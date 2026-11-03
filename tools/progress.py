"""Regenerate progress.md for all provinces under data/.
Usage: python3 tools/progress.py <省份> "<当前轮次>" "<下一步>"   (state kept in tools/progress_state.json)"""
import csv, glob, json, os, sys
R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SF = os.path.join(R, 'tools', 'progress_state.json')
state = json.load(open(SF, encoding='utf-8')) if os.path.exists(SF) else {}
if len(sys.argv) >= 4:
    state[sys.argv[1]] = {'轮次': sys.argv[2], '下一步': sys.argv[3]}
    json.dump(state, open(SF, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
TODAY = '2026-11-03'
provs = []
for d in sorted(glob.glob(os.path.join(R, 'data', '*', 'gdp_2025.csv')), key=lambda p: list(state).index(os.path.basename(os.path.dirname(p))) if os.path.basename(os.path.dirname(p)) in state else 99):
    name = os.path.basename(os.path.dirname(d))
    base = list(csv.DictReader(open(os.path.join(os.path.dirname(d), 'base.csv'), encoding='utf-8')))
    g = {r['县级行政区']: r for r in csv.DictReader(open(d, encoding='utf-8'))}
    provs.append((name, base, g))
def st(rs):
    v = sum(1 for r in rs if r['GDP_亿元']); b = sum(1 for r in rs if r['核验状态'] == '空白终态'); p = sum(1 for r in rs if r['核验状态'] == '待补缺')
    return v, b, p
L = ['# 进度（新会话先读本文件）', '', f'更新：{TODAY}', '', '## 各省总表', '',
     '| 省份 | 底表数 | 已完成 | 精确值 | 空白终态 | 待补缺 | 当前轮次 | 更新日期 |', '|---|---|---|---|---|---|---|---|']
for name, base, g in provs:
    v, b, p = st(list(g.values()))
    L.append(f"| {name} | {len(base)} | {v+b} | {v} | {b} | {p} | {state.get(name,{}).get('轮次','')} | {TODAY} |")
for name, base, g in provs:
    L += ['', f'## {name}分批', '', '| 批次 | 区县数 | 已完成 | 精确值 | 空白终态 | 待补缺区县 |', '|---|---|---|---|---|---|']
    order = list(dict.fromkeys(r['批次'] for r in base))
    for bt in order:
        ns = [r['县级行政区'] for r in base if r['批次'] == bt]
        v, b, p = st([g[n] for n in ns])
        pend = '、'.join(n for n in ns if g[n]['核验状态'] == '待补缺') or '—'
        L.append(f'| {bt} | {len(ns)} | {v+b} | {v} | {b} | {pend} |')
L += ['', '## 核验方式（2026-11 会话）', '',
      '本会话的网络策略拦截所有 `*.gov.cn`，WebFetch 和 curl 都打不开。按用户 2026-11-03 指示，改用内置 WebSearch：', '',
      '- 检索时限定官方域名（allowed_domains），由搜索服务读取官方页面原文。',
      '- 每个数值至少用两条独立检索确认：同一官方 URL，原文同时含该数值和"2025年"。',
      '- 每行备注标明"WebSearch 服务端读取，待直连复核"。',
      '- 网络放开后，应对全部已核验行用 WebFetch 直连复核一次。',
      '- WebSearch 每轮约 200 次、所有子代理共用；额度用完时先 commit 断点，下一轮接着做。',
      '- 汇总 Excel：`python3 tools/build_xlsx.py "<版本>"` 重新生成 output/县级GDP_2025_汇总.xlsx。',
      '', '## 下一步', '']
for name, _, _ in provs:
    L.append(f"- **{name}**：{state.get(name,{}).get('下一步','')}")
open(os.path.join(R, 'progress.md'), 'w', encoding='utf-8').write('\n'.join(L) + '\n')
print([(n, st(list(g.values()))) for n, _, g in provs])
