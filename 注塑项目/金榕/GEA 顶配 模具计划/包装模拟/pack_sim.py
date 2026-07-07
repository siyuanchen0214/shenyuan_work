#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GEA 顶棚框架 · 包装摆放模拟优化器  (pack_sim)
================================================
暴力枚举「零件怎么摆进纸箱、纸箱怎么码进货车」，找「纸箱料费+运费」综合最省的方案。
分装(每种件单独装箱) 和 混装(前后左右同箱) 都会跑并对比。

两层模型：
    零件(4种, 6种朝向) ─► 纸箱(箱型由程序优化) ─► 货车车厢(体积+限重)
    总成本 = 纸箱料费(∑ 箱外表面积 × 纸板单价) + 运费(需要几趟车 × 每趟车价)

参数全在 config.py（可调；带 ★待确认 的是占位值，换成真实数字结果才有意义）。
运行：  python3 pack_sim.py   —— 结果同时打印 + 写入 outputs/ 的带时间戳报告。
依赖：py3dbp（混装用）；缺了会自动跳过混装。
"""
from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from itertools import permutations
from math import ceil, floor
from pathlib import Path
from typing import Optional
import config as C


def output_dir() -> Path:
    """报告输出目录：config.OUTPUT_DIR 为空则用脚本所在文件夹；
    相对路径按脚本目录解析，也接受绝对路径。"""
    base = Path(__file__).parent
    d = getattr(C, "OUTPUT_DIR", "") or ""
    if not d:
        return base
    p = Path(d)
    return p if p.is_absolute() else base / p


# ============================ 数据结构 ============================
@dataclass
class CartonPlan:
    label: str
    inner_cm: tuple
    outer_cm: tuple
    contents: dict           # {零件名: 数量}
    weight_kg: Optional[float]

    @property
    def outer_volume_m3(self):
        L, W, H = self.outer_cm
        return (L/100)*(W/100)*(H/100)

    @property
    def material_yuan(self):
        L, W, H = (d/100 for d in self.outer_cm)
        surface_m2 = 2*(L*W + L*H + W*H)
        return surface_m2 * C.CARDBOARD_YUAN_PER_M2

    @property
    def parts_total(self):
        return sum(self.contents.values())


# ============================ 工具 ============================
def orientations(L, W, H):
    seen = set()
    for p in permutations((L, W, H)):
        if p not in seen:
            seen.add(p); yield p

def carton_outer(inner):
    return tuple(round(d + 2*C.CARTON_WALL_CM, 2) for d in inner)

def fmt_cm(t):
    return "×".join(f"{v:.1f}" for v in t)


# ============================ 分装：网格暴力 ============================
def best_separate_cartons():
    result = {}
    for name, L, W, H, w in C.PARTS:
        best = None
        for (px, py, pz) in orientations(L, W, H):
            for nx in range(1, C.MAX_PARTS_PER_CARTON + 1):
                inner_L = nx*px + 2*C.PART_CLEARANCE_CM
                if inner_L + 2*C.CARTON_WALL_CM > C.MAX_CARTON_EDGE_CM: break
                for ny in range(1, C.MAX_PARTS_PER_CARTON + 1):
                    inner_W = ny*py + 2*C.PART_CLEARANCE_CM
                    if inner_W + 2*C.CARTON_WALL_CM > C.MAX_CARTON_EDGE_CM: break
                    for nz in range(1, C.MAX_PARTS_PER_CARTON + 1):
                        n = nx*ny*nz
                        if n > C.MAX_PARTS_PER_CARTON: break
                        inner_H = nz*pz + 2*C.PART_CLEARANCE_CM
                        if inner_H + 2*C.CARTON_WALL_CM > C.MAX_CARTON_EDGE_CM: break
                        inner = (round(inner_L,2), round(inner_W,2), round(inner_H,2))
                        outer = carton_outer(inner)
                        if any(o > t for o, t in zip(sorted(outer), sorted(C.TRUCK_INNER_CM))):
                            continue
                        wt = (w or 0)*n if w is not None else None
                        plan = CartonPlan(f"{name}×{n} ({nx}×{ny}×{nz})",
                                          inner, outer, {name: n}, wt)
                        score = plan.material_yuan / n
                        if best is None or score < best[0]:
                            best = (score, plan)
        result[name] = best[1]
    return result


# ============================ 混装：py3dbp ============================
def best_mixed_carton():
    try:
        from py3dbp import Packer, Bin, Item
    except ImportError:
        return None, "未安装 py3dbp，跳过混装（pip install py3dbp）"
    best = None
    max_len = max(L for _, L, W, H, w in C.PARTS) + 2*C.PART_CLEARANCE_CM
    for k in range(1, C.MAX_PARTS_PER_CARTON // len(C.PARTS) + 1):
        if k*len(C.PARTS) > C.MAX_PARTS_PER_CARTON: break
        packer = Packer()
        # 长边卡在最长件，逼它在横截面码成紧凑一捆，而非首尾排长条
        packer.add_bin(Bin("carton", max_len, C.MAX_CARTON_EDGE_CM, C.MAX_CARTON_EDGE_CM, 1e9))
        total = 0
        for name, L, W, H, w in C.PARTS:
            for i in range(k):
                packer.add_item(Item(f"{name}#{i}", L, W, H, (w or 0))); total += 1
        packer.pack(bigger_first=True, distribute_items=False, number_of_decimals=1)
        fitted = packer.bins[0].items
        if len(fitted) < total: continue
        maxx = maxy = maxz = 0.0
        for it in fitted:
            dx, dy, dz = it.get_dimension(); x, y, z = it.position
            maxx = max(maxx, float(x)+float(dx))
            maxy = max(maxy, float(y)+float(dy))
            maxz = max(maxz, float(z)+float(dz))
        inner = (round(maxx+2*C.PART_CLEARANCE_CM,1), round(maxy+2*C.PART_CLEARANCE_CM,1),
                 round(maxz+2*C.PART_CLEARANCE_CM,1))
        outer = carton_outer(inner)
        if any(o > t for o, t in zip(sorted(outer), sorted(C.TRUCK_INNER_CM))): continue
        wt = sum(w for *_, w in C.PARTS)*k if all(w is not None for *_, w in C.PARTS) else None
        plan = CartonPlan(f"混装 ×{k}套", inner, outer,
                          {name: k for name, *_ in C.PARTS}, wt)
        score = plan.material_yuan / plan.parts_total
        if best is None or score < best[0]:
            best = (score, plan)
    return (best[1] if best else None), None


# ============================ 车厢装载 + 成本 ============================
def cartons_per_truck(plan):
    usable = (C.TRUCK_INNER_CM[0]/100)*(C.TRUCK_INNER_CM[1]/100)*(C.TRUCK_INNER_CM[2]/100)*C.TRUCK_FILL_EFF
    by_vol = floor(usable / plan.outer_volume_m3)
    if plan.weight_kg:
        return max(1, min(by_vol, floor(C.TRUCK_PAYLOAD_KG / plan.weight_kg)))
    return max(1, by_vol)

def evaluate(cartons, sets):
    need = {name: sets for name, *_ in C.PARTS}
    total_material = 0.0; total_boxes = 0; truck_slots = 0.0; detail = []
    for plan in cartons:
        boxes = max(ceil(need[n]/c) for n, c in plan.contents.items())
        for n, c in plan.contents.items(): need[n] -= boxes*c
        total_material += boxes*plan.material_yuan
        total_boxes += boxes
        cpt = cartons_per_truck(plan)
        truck_slots += boxes/cpt
        detail.append((plan, boxes, cpt))
    trucks = ceil(truck_slots)
    return dict(total=total_material + trucks*C.TRUCK_COST_YUAN, material=total_material,
                freight=trucks*C.TRUCK_COST_YUAN, trucks=trucks, boxes=total_boxes,
                detail=detail, leftover={n: v for n, v in need.items() if v > 0})


# ============================ 报告 ============================
def build_report():
    out = []
    def line(s=""): out.append(s)

    line(f"# 包装模拟报告 · {datetime.now():%Y-%m-%d %H:%M}")
    line()
    if C.PENDING:
        line("> ⚠️ **仍在用占位值的参数（结果仅供参考，填真实值后重跑）：**")
        for p in C.PENDING: line(f"> - {p}")
        line()
    line(f"**订单** {C.SETS_TO_SHIP} 套（每套4件） | **车厢** {fmt_cm(C.TRUCK_INNER_CM)}cm "
         f"| **纸板** {C.CARDBOARD_YUAN_PER_M2}元/m² | **运费** {C.TRUCK_COST_YUAN}元/趟")
    line()

    def section(title, cartons, res):
        line(f"## {title}")
        for plan, boxes, cpt in res["detail"]:
            cont = ", ".join(f"{n}×{c}" for n, c in plan.contents.items())
            line(f"- **{plan.label}** — 箱内 {fmt_cm(plan.inner_cm)} / 外 {fmt_cm(plan.outer_cm)} cm")
            line(f"    - 装[{cont}] | 单箱料费 {plan.material_yuan:.1f}元 | 需 {boxes} 箱 | 每趟车 {cpt} 箱")
        if res["leftover"]: line(f"    - ⚠️ 未覆盖: {res['leftover']}")
        line(f"- **汇总**：纸箱 {res['boxes']} 个 · 车 {res['trucks']} 趟 · "
             f"料费 {res['material']:.0f} + 运费 {res['freight']:.0f} = "
             f"**总成本 {res['total']:.0f}元**（每套 {res['total']/C.SETS_TO_SHIP:.2f}元）")
        line()

    resA = evaluate(list(best_separate_cartons().values()), C.SETS_TO_SHIP)
    section("方案A · 分装（每种件单独装箱）", None, resA)

    mix, warn = best_mixed_carton()
    resB = None
    if mix:
        resB = evaluate([mix], C.SETS_TO_SHIP)
        section("方案B · 混装（前后左右同箱）", None, resB)
    elif warn:
        line(f"## 方案B · 混装\n- {warn}\n")

    line("## 结论")
    if resB and resB["total"] < resA["total"]:
        d = resA["total"]-resB["total"]
        line(f"**混装更省** — 总 {resB['total']:.0f}元，比分装省 {d:.0f}元（{d/resA['total']*100:.1f}%）")
    else:
        d = (resB["total"]-resA["total"]) if resB else 0
        line(f"**分装更省** — 总 {resA['total']:.0f}元" + (f"，比混装省 {d:.0f}元" if resB else ""))
    return "\n".join(out)


def main():
    report = build_report()
    print(report)
    out = output_dir()
    out.mkdir(parents=True, exist_ok=True)
    fp = out / f"pack_report_{datetime.now():%Y%m%d_%H%M%S}.md"
    fp.write_text(report, encoding="utf-8")
    print(f"\n[已写出报告] {fp}")


if __name__ == "__main__":
    main()
