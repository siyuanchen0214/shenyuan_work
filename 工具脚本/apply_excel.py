#!/usr/bin/env python3
# 把 trans_excel.T 回填到镜面覆盖塑料件2 的所有 Excel(同格换行加中文)。
# .xlsx/.xlsm 原地改; .xls 转成同名 .xlsx 双语(原.xls保留)。报告覆盖率。
import sys, os, re, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, "/Users/chensiyuan5/Desktop/Github/shenyuan_work/工具脚本")
import bilingual as B
from trans_excel import T

BASE = "/Users/chensiyuan5/Desktop/Github/shenyuan_work/注塑项目/金榕/未报价 镜面覆盖塑料件 2"
KM = BASE + "/｣ｨKM -Mirror Protector -2026｣ｩ---Technical Solution"
EJ = BASE + "/EJ Door Bags and Straps---Technical Solution"

xlsx_files = [
    (EJ + "/EJ Door bags DVP.xlsx", False),
    (KM + "/Mirror protector - DVP.xlsx", False),
    (KM + "/PIS Format.xlsm", True),
    (EJ + "/PIS Format (1).xlsm", True),
]
xls_files = [
    EJ + "/ADP RR-Mopar Interior .xls",
    EJ + "/ATAR (2).xls",
    EJ + "/EDD template.xls",
]

for p, vba in xlsx_files:
    B.annotate_xlsx(p, p, T, keep_vba=vba)
    print("OK xlsx:", os.path.basename(p))

for p in xls_files:
    out = p[:-4] + ".xlsx"
    B.annotate_xls_to_xlsx(p, out, T)
    print("OK xls->xlsx:", os.path.basename(out))

# 覆盖率统计
import openpyxl, xlrd
def keep(s):
    s=str(s).strip()
    return bool(s) and bool(re.search(r'[A-Za-z]',s)) and not re.fullmatch(r'[-+]?\d[\d,.\s%/]*',s)
hit=miss=0; missed=set()
def chk(v):
    global hit,miss
    if keep(v):
        if str(v).strip() in T: hit+=1
        else: miss+=1; missed.add(str(v).strip())
for p,_ in xlsx_files:
    wb=openpyxl.load_workbook(p,data_only=True,read_only=True)
    for ws in wb.worksheets:
        if ws.title=="FontSheet": continue
        n=0
        for r in ws.iter_rows(values_only=True):
            n+=1
            if n>3000:break
            for c in r:
                # 注意已回填,英文里现在带【中】;只统计未带中文的英文
                if c and "【中】" not in str(c): chk(c)
for p in xls_files:
    wb=xlrd.open_workbook(p)
    for ws in wb.sheets():
        for r in range(min(ws.nrows,3000)):
            for c in range(ws.ncols): chk(ws.cell_value(r,c))
print(f"\n覆盖: 命中{hit} 未译{miss}")
open("/tmp/missed.txt","w").write("\n".join(sorted(missed,key=lambda x:(len(x),x))))
print("未译清单 -> /tmp/missed.txt")
