"""
穴数优化 Simulation
用法：给定单个产品参数，自动找出最优穴数方案
"""
import math

# 可用机台（吨位→[设备价万元, 额定功率KW]），待补充完整
MACHINES = {
    # 价格锚点：1300T=130万封顶，其余只能更低（用户确认）；价格仍为估值，待真实参数校准
    60:   [12,   22],
    100:  [16,   37],
    120:  [19,   45],
    140:  [22,   55],
    250:  [40,   90],
    280:  [44,   100],
    320:  [50,   110],
    380:  [58,   130],
    450:  [68,   150],
    650:  [85,   200],
    850:  [105,  280],
    1300: [130,  400],
}

# 模具成本倍数（相对单穴）
MOLD_MULTIPLIER = {1: 1.0, 2: 1.4, 4: 1.8, 6: 2.2, 8: 2.5, 12: 3.0}

def select_machine(required_tonnage):
    """向上取最近可用机台，返回 (吨位, 价格万, 功率KW) 或 None"""
    for t in sorted(MACHINES.keys()):
        if t >= required_tonnage:
            price_wan, power = MACHINES[t]
            return t, price_wan * 10000, power
    return None

def optimize_cavities(
    proj_area_cm2,      # 单穴投影面积 cm²
    cavity_pressure,    # 型腔压力 kg/cm²（PP≈300, ABS/PC≈400）
    cycle_s,            # 成型周期 s
    net_weight_kg,      # 净用量 kg/件（含浇口）
    mat_price,          # 材料单价 元/kg
    mold_base_price,    # 单穴模具基础价 元
    total_qty,          # 总产量 件
    operators=2,        # 操作人数
    oee=0.75,           # 稼动率
    labor_rate=24,      # 元/H
    energy_util=0.45,   # 能耗利用率
    energy_price=1.0,   # 电价 元/KWh
    safety_factor=1.2,  # 锁模力安全系数
    waste_rate=0.01,    # 废品率
):
    results = []

    for n_cavities, mold_mult in sorted(MOLD_MULTIPLIER.items()):
        # 1. 锁模力
        required_t = proj_area_cm2 * n_cavities * cavity_pressure / 1000 * safety_factor
        machine = select_machine(required_t)
        if machine is None:
            continue
        tonnage, machine_price, power = machine

        # 2. 模具摊销
        mold_total = mold_base_price * mold_mult
        mold_amort = mold_total / total_qty

        # 3. 材料费
        mat_cost = net_weight_kg * mat_price * (1 + waste_rate)

        # 4. 人工费（per piece）
        labor = labor_rate * (cycle_s / 3600) * operators / oee / n_cavities

        # 5. 折旧（per piece）
        depr = machine_price * 0.9 / 66000 * (cycle_s / 3600) / oee / n_cavities

        # 6. 能耗（per piece）
        energy = power * energy_util * (14/13) * (cycle_s / 3600) / n_cavities * energy_price

        total_cost = mat_cost + mold_amort + labor + depr + energy

        results.append({
            'cavities': n_cavities,
            'tonnage':  tonnage,
            'mold_total_wan': mold_total / 10000,
            'mat_cost': mat_cost,
            'mold_amort': mold_amort,
            'labor': labor,
            'depr': depr,
            'energy': energy,
            'total': total_cost,
        })

    results.sort(key=lambda x: x['total'])

    print(f"{'穴数':>4} {'机台':>6} {'模具万':>6} {'材料':>7} {'模摊':>7} {'人工':>7} {'折旧':>7} {'能耗':>7} {'总成本':>8}")
    print('-' * 65)
    best = results[0]['total']
    for r in sorted(results, key=lambda x: x['cavities']):
        marker = ' ★' if r['total'] == best else ''
        print(f"{r['cavities']:>4}穴 {r['tonnage']:>4}T  {r['mold_total_wan']:>5.1f}万"
              f"  {r['mat_cost']:>6.4f}  {r['mold_amort']:>6.4f}  {r['labor']:>6.4f}"
              f"  {r['depr']:>6.4f}  {r['energy']:>6.4f}  {r['total']:>7.4f}{marker}")

    print(f"\n推荐方案：{results[0]['cavities']}穴 / {results[0]['tonnage']}T，总成本 ¥{results[0]['total']:.4f}/件")
    return results


if __name__ == '__main__':
    print("=== Lucid 地板舱盖 穴数优化 ===")
    optimize_cavities(
        proj_area_cm2    = 122.5,
        cavity_pressure  = 300,
        cycle_s          = 10,
        net_weight_kg    = 0.032,
        mat_price        = 9.1,
        mold_base_price  = 70000 / 1.8,  # 正确答案4穴7万 → 单穴基础价约3.9万
        total_qty        = 875000,
        operators        = 2,
    )
