"""Build output/县级GDP_2025_汇总.xlsx from data/<省份>/gdp_2025.csv (CLAUDE.md §7)."""
import csv, glob, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = '/home/user/county-gdp-2025'
OUT = os.path.join(ROOT, 'output', '县级GDP_2025_汇总.xlsx')
VERSION = sys.argv[1] if len(sys.argv) > 1 else 'v2026-11-03'
HDR = ['省级行政区', '地级地区', '县级行政区', '行政类型', '统计年份', 'GDP_亿元', '核验状态', '来源类型',
       '来源标题', '发布机构', '来源URL', '原文摘录', '已查渠道', '排除的候选值', '核验日期', '备注']
WIDTH = {'省级行政区': 10, '地级地区': 10, '县级行政区': 20, '行政类型': 9, '统计年份': 8, 'GDP_亿元': 11,
         '核验状态': 10, '来源类型': 12, '来源标题': 36, '发布机构': 24, '来源URL': 50, '原文摘录': 48,
         '已查渠道': 50, '排除的候选值': 40, '核验日期': 11, '备注': 60}
F = 'Arial'
font = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
hfont = Font(name=F, size=10, bold=True, color='FFFFFF')
hfill = PatternFill('solid', fgColor='305496')
red = PatternFill('solid', fgColor='FFC7CE')       # 空白行浅红
link = Font(name=F, size=10, color='0563C1', underline='single')
thin = Side(style='thin', color='BFBFBF')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical='top')

# ---- load data ----
provs = []
for p in sorted(glob.glob(os.path.join(ROOT, 'data', '*', 'gdp_2025.csv'))):
    rows = list(csv.DictReader(open(p, encoding='utf-8')))
    assert list(rows[0].keys()) == HDR, p
    base = os.path.join(os.path.dirname(p), 'base.csv')
    nbase = len(list(csv.DictReader(open(base, encoding='utf-8')))) if os.path.exists(base) else len(rows)
    assert nbase == len(rows), (p, nbase, len(rows))
    assert len({r['县级行政区'] for r in rows}) == len(rows), ('dup', p)
    provs.append((rows[0]['省级行政区'], rows))

wb = Workbook()

def detail_sheet(ws, rows):
    ws.append(HDR)
    for c in ws[1]:
        c.font, c.fill, c.border = hfont, hfill, box
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    gi = HDR.index('GDP_亿元'); ui = HDR.index('来源URL')
    for r in rows:
        vals = [r[h] for h in HDR]
        vals[HDR.index('统计年份')] = int(vals[HDR.index('统计年份')])
        vals[gi] = float(vals[gi]) if vals[gi] else None
        ws.append(vals)
        i = ws.max_row
        blank = vals[gi] is None
        for j, c in enumerate(ws[i]):
            c.font, c.border, c.alignment = font, box, wrap
            if blank:
                c.fill = red
        u = r['来源URL']
        if u:
            c = ws.cell(row=i, column=ui + 1)
            c.hyperlink = u
            c.font = link
    for j, h in enumerate(HDR, 1):
        ws.column_dimensions[get_column_letter(j)].width = WIDTH[h]
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HDR))}{ws.max_row}"

# ---- 总览 ----
ov = wb.active
ov.title = '总览'
ov['A1'] = '县级行政区 2025 年 GDP 官方数据汇总 · 总览'
ov['A1'].font = Font(name=F, size=14, bold=True)
ov['A2'] = f'版本：{VERSION}　　数量与合计均由公式从"全国汇总表"计算'
ov['A2'].font = Font(name=F, size=9, italic=True, color='595959')
cols = ['省份', '底表数', '精确值数', '空白终态数', '待补缺数', '覆盖率', 'GDP合计（亿元）', '版本']
ov.append([])
ov.append(cols)
hr = ov.max_row
for c in ov[hr]:
    c.font, c.fill, c.border = hfont, hfill, box
    c.alignment = Alignment(horizontal='center', vertical='center')
R = "'全国汇总表'"
first = hr + 1
for name, rows in provs:
    i = ov.max_row + 1
    ov.cell(row=i, column=1, value=name)
    ov.cell(row=i, column=2, value=f'=COUNTIFS({R}!$A:$A,$A{i})')
    ov.cell(row=i, column=3, value=f'=COUNTIFS({R}!$A:$A,$A{i},{R}!$F:$F,"<>")')
    ov.cell(row=i, column=4, value=f'=COUNTIFS({R}!$A:$A,$A{i},{R}!$G:$G,"空白终态")')
    ov.cell(row=i, column=5, value=f'=COUNTIFS({R}!$A:$A,$A{i},{R}!$G:$G,"待补缺")')
    ov.cell(row=i, column=6, value=f'=IF(B{i}=0,0,C{i}/B{i})')
    ov.cell(row=i, column=7, value=f'=SUMIFS({R}!$F:$F,{R}!$A:$A,$A{i})')
    ov.cell(row=i, column=8, value=VERSION)
last = ov.max_row
t = last + 1
ov.cell(row=t, column=1, value='合计')
for col in 'BCDEG':
    ov[f'{col}{t}'] = f'=SUM({col}{first}:{col}{last})'
ov[f'F{t}'] = f'=IF(B{t}=0,0,C{t}/B{t})'
for row in ov.iter_rows(min_row=first, max_row=t, max_col=len(cols)):
    for c in row:
        c.font, c.border = font, box
for c in ov[t][:len(cols)]:
    c.font = bold
for i in range(first, t + 1):
    ov[f'F{i}'].number_format = '0.0%'
    ov[f'G{i}'].number_format = '#,##0.00'
for j, w in enumerate([12, 9, 10, 11, 9, 9, 16, 34], 1):
    ov.column_dimensions[get_column_letter(j)].width = w
n = t + 2
ov.cell(row=n, column=1, value='说明：浅红色填充 = GDP 空白行；GDP合计只加总有精确值的区县，空白区县不计入，不能视为全省合计。').font = Font(name=F, size=9, color='595959')

# ---- 全国汇总表 + 各省 ----
allrows = [r for _, rows in provs for r in rows]
detail_sheet(wb.create_sheet('全国汇总表'), allrows)
for name, rows in provs:
    detail_sheet(wb.create_sheet(name), rows)

# ---- 口径说明 ----
ks = wb.create_sheet('口径说明')
notes = open(os.path.join(ROOT, 'tools', 'koujing.txt'), encoding='utf-8').read().strip().split('\n')
for line in notes:
    ks.append([line])
    c = ks.cell(row=ks.max_row, column=1)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    c.font = Font(name=F, size=12, bold=True) if line.startswith('#') else font
    if line.startswith('#'):
        c.value = line.lstrip('# ')
ks.column_dimensions['A'].width = 120

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print('saved', OUT, [(n, len(r)) for n, r in provs])
