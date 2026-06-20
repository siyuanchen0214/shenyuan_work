"""
STP/STEP 体积计算 — 逐实体精确体积，用于注塑成品克重。

⚠️ 用 gmsh/OCC 直接读 STEP 实体的精确体积（OpenCASCADE GProp），不走网格。
   不要再用 cascadio→OBJ→trimesh 的老路：OBJ 转换会按材料把实体合并、破坏水密性，
   导致体积聚合错误（P06大灯实测：网格法把 161cc 误算成 271cc）。STEP 本身是标准
   可精确读取的，per-solid 体积应当直接出。

用法：
    python3 stp_volume.py <文件或文件夹> [--density 1.20] [--pair]

输出每个实体的精确体积(cc) + 包围盒 + 质心；给 --density(g/cc) 时一并算克重。
    --pair  按 Y 质心识别左右对称件(L/R)，把成对实体归为同一零件型。
依赖：gmsh（pip install gmsh）。装配体内每个 solid 独立列出，不做合并。
"""
import sys
import os
import glob
import argparse


def solids_of(stp_path: str):
    """返回 [(tag, vol_cc, (cx,cy,cz), (ex,ey,ez)), ...]，按体积降序。用 OCC 精确体积。"""
    import gmsh
    gmsh.initialize()
    gmsh.option.setNumber("General.Terminal", 0)
    try:
        gmsh.model.add("m")
        gmsh.model.occ.importShapes(stp_path)
        gmsh.model.occ.synchronize()
        out = []
        for dim, tag in gmsh.model.getEntities(3):  # 3D = solids
            vol = gmsh.model.occ.getMass(dim, tag) / 1000.0          # mm^3 -> cc
            cx, cy, cz = gmsh.model.occ.getCenterOfMass(dim, tag)
            bb = gmsh.model.getBoundingBox(dim, tag)
            ext = (bb[3] - bb[0], bb[4] - bb[1], bb[5] - bb[2])
            out.append((tag, vol, (cx, cy, cz), ext))
        return sorted(out, key=lambda x: -x[1])
    finally:
        gmsh.finalize()


def pair_lr(solids, tol_cc=0.5):
    """按体积近似 + Y 质心异号 配 L/R。返回 [(vol_cc, count, [tags])]，每个零件型一行。"""
    used = set()
    groups = []
    for i, (tag, v, c, e) in enumerate(solids):
        if tag in used:
            continue
        mates = [tag]
        for tag2, v2, c2, e2 in solids[i + 1:]:
            if tag2 in used:
                continue
            if abs(v - v2) <= tol_cc and c[1] * c2[1] < 0:   # 体积近似且 Y 异号
                mates.append(tag2)
                used.add(tag2)
                break
        used.add(tag)
        groups.append((v, len(mates), mates))
    return groups


def analyze(stp_path, density=None, do_pair=False):
    solids = solids_of(stp_path)
    print(f"==== {os.path.basename(stp_path)}  ({len(solids)} solid) ====")
    total = 0.0
    for tag, v, c, e in solids:
        total += v
        line = f"  solid#{tag}: {v:8.2f} cc   bbox {e[0]:.0f}x{e[1]:.0f}x{e[2]:.0f} mm   Y={c[1]:.0f}"
        if density:
            line += f"   {v * density:.1f} g"
        print(line)
    print(f"  ---- 合计 {total:.1f} cc" + (f" = {total * density:.0f} g" if density else ""))
    if do_pair:
        print("  -- 按 L/R 配对后的零件型 --")
        for v, n, tags in pair_lr(solids):
            g = f"{v * density:.1f} g" if density else ""
            print(f"     {v:8.2f} cc  ×{n}件(L/R)  tags={tags}  {g}")
    return solids


def main():
    ap = argparse.ArgumentParser(description="STP/STEP 逐实体精确体积 (gmsh/OCC)")
    ap.add_argument("path", help="STP 文件或含 STP 的文件夹")
    ap.add_argument("--density", type=float, help="密度 g/cc，给了就算克重")
    ap.add_argument("--pair", action="store_true", help="按 Y 质心识别 L/R 对称件")
    args = ap.parse_args()

    if os.path.isdir(args.path):
        files = sorted(glob.glob(os.path.join(args.path, "**", "*.stp"), recursive=True) +
                       glob.glob(os.path.join(args.path, "**", "*.step"), recursive=True))
    else:
        files = [args.path]
    if not files:
        print("未找到 STP 文件")
        return
    for f in files:
        analyze(f, args.density, args.pair)


if __name__ == "__main__":
    main()
