"""Usage: python3 tools/apply.py <省份> <records.json>  — merge verified records into data/<省份>/gdp_2025.csv by 县级行政区."""
import csv, json, sys
import os
CSV = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', sys.argv[1], 'gdp_2025.csv')
STATUS = {'已核验', '空白终态', '待补缺'}
TYPES = {'', '统计公报', '统一核算', '全年统计表', '统计年鉴', '计划执行报告', '政府工作报告', '规划回顾', '县情页面', '其他官方'}
recs = json.load(open(sys.argv[2], encoding='utf-8'))
rows = list(csv.DictReader(open(CSV, encoding='utf-8')))
hdr = list(rows[0].keys())
idx = {r['县级行政区']: r for r in rows}
for rec in recs:
    name = rec['县级行政区']
    assert name in idx, name
    row = idx[name]
    for k, v in rec.items():
        assert k in hdr, k
        row[k] = '' if v is None else str(v)
    row['核验日期'] = row['核验日期'] or '2026-11-03'
    SUF = '核验方式：WebSearch 服务端读取官方页面原文（本会话无法直连 gov.cn），待直连复核。'
    if row['GDP_亿元'] and SUF not in row['备注']:
        row['备注'] = (row['备注'] + ' ' + SUF).strip()
    assert row['核验状态'] in STATUS, (name, row['核验状态'])
    assert row['来源类型'] in TYPES, (name, row['来源类型'])
    if row['GDP_亿元']:
        float(row['GDP_亿元'])
        assert row['核验状态'] == '已核验' and row['来源URL'] and row['原文摘录'], name
        assert len(row['原文摘录']) <= 60, (name, len(row['原文摘录']))
    else:
        assert row['核验状态'] != '已核验', name
    if row['核验状态'] == '空白终态':
        assert row['已查渠道'], name
with open(CSV, 'w', encoding='utf-8', newline='') as f:
    w = csv.DictWriter(f, fieldnames=hdr); w.writeheader(); w.writerows(rows)
from collections import Counter
print('updated', len(recs), Counter(r['核验状态'] for r in rows), 'with value:', sum(1 for r in rows if r['GDP_亿元']))
