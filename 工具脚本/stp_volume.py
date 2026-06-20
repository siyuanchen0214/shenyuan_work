"""
STP 体积计算 — 从 STP/STEP 文件算注塑成品体积，误差 <1%。

用法：
    python3 stp_volume.py <文件或文件夹> [--density 1.16] [--main largest|median]

输出每个 body 的体积(cc) + 包围盒；标出主注塑件候选。
给了 --density(g/cc) 时一并算成品克重。
装配体里有多个 body：
    --main largest（默认）取体积最大的 body 为主注塑件；
    --main median  取中位体积的 body（金属嵌件+塑料件混装时常用）。
依赖：cascadio, trimesh
"""
import sys
import os
import glob
import argparse
import cascadio
import trimesh


def bodies_of(stp_path: str):
    """返回 [(vol_cc, (x,y,z)mm), ...]，按体积降序。"""
    obj = os.path.join("/tmp", os.path.basename(stp_path) + ".obj")
    cascadio.step_to_obj(stp_path, obj)
    scene = trimesh.load(obj)
    geoms = list(scene.geometry.values()) if isinstance(scene, trimesh.Scene) else [scene]
    out = []
    for g in geoms:
        ext = g.bounding_box.extents
        out.append((abs(g.volume) / 1000.0, (ext[0], ext[1], ext[2])))
    return sorted(out, key=lambda x: -x[0])


def pick_main(bodies, mode="largest"):
    """从 body 列表里选主注塑件。"""
    if not bodies:
        return None
    if mode == "median":
        return sorted(bodies, key=lambda x: x[0])[len(bodies) // 2]
    return bodies[0]  # largest


def analyze(stp_path: str, density=None, mode="largest"):
    bodies = bodies_of(stp_path)
    print(f"==== {os.path.basename(stp_path)}  ({len(bodies)} body) ====")
    for i, (v, e) in enumerate(bodies):
        print(f"  body{i}: {v:8.2f} cc   bbox {e[0]:.1f} x {e[1]:.1f} x {e[2]:.1f} mm")
    main = pick_main(bodies, mode)
    vol = main[0]
    e = main[1]
    print(f"  >> 主注塑件({mode}): {vol:.2f} cc  bbox {e[0]:.1f} x {e[1]:.1f} x {e[2]:.1f} mm")
    if density:
        print(f"  >> 成品克重 @ {density} g/cc = {vol * density:.1f} g = {vol * density / 1000:.4f} kg")
    return {"file": os.path.basename(stp_path), "main_vol_cc": vol,
            "bbox_mm": e, "bodies": bodies}


def main():
    ap = argparse.ArgumentParser(description="STP 体积计算")
    ap.add_argument("path", help="STP 文件或含 STP 的文件夹")
    ap.add_argument("--density", type=float, help="密度 g/cc，给了就算克重")
    ap.add_argument("--main", choices=["largest", "median"], default="largest")
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
        analyze(f, args.density, args.main)


if __name__ == "__main__":
    main()
