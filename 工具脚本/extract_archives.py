"""
通用解压脚本 — 递归解压询价资料里的 zip，正确处理中文（GBK）文件名。

为什么需要它：
- 苹果自带 unzip 不支持 -O 指定编码，中文文件名会乱码。
- 询价 zip 常常嵌套 zip（外层一个、每个零件一个）。

用法：
    python3 extract_archives.py <文件夹或zip路径> [-o 输出目录]
不指定 -o 时，解压到 <输入同级>/_unzip。
默认递归解压嵌套 zip。
"""
import sys
import os
import zipfile
import argparse


def _decode(name: str) -> str:
    """把 zip 里 cp437 存储的文件名按 GBK / UTF-8 还原。"""
    for enc in ("gbk", "utf-8"):
        try:
            return name.encode("cp437").decode(enc)
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return name


def extract_zip(zip_path: str, out_dir: str) -> list:
    """解压单个 zip 到 out_dir，返回解出的文件绝对路径列表。"""
    extracted = []
    with zipfile.ZipFile(zip_path) as z:
        for info in z.infolist():
            name = _decode(info.filename)
            target = os.path.join(out_dir, name)
            if info.is_dir():
                os.makedirs(target, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(target) or out_dir, exist_ok=True)
            with z.open(info) as src, open(target, "wb") as dst:
                dst.write(src.read())
            extracted.append(target)
    return extracted


def extract_recursive(path: str, out_dir: str, depth: int = 0) -> list:
    """解压 path（zip 或含 zip 的文件夹），递归处理嵌套 zip。"""
    os.makedirs(out_dir, exist_ok=True)
    all_files = []

    # 收集本层要解压的 zip
    zips = []
    if os.path.isfile(path) and path.lower().endswith(".zip"):
        zips = [path]
    elif os.path.isdir(path):
        zips = [os.path.join(path, f) for f in sorted(os.listdir(path))
                if f.lower().endswith(".zip")]

    for zp in zips:
        sub_out = os.path.join(out_dir, os.path.splitext(os.path.basename(zp))[0])
        files = extract_zip(zp, sub_out)
        print("  " * depth + f"[解压] {os.path.basename(zp)} -> {len(files)} 个文件")
        for f in files:
            if f.lower().endswith(".zip"):
                all_files += extract_recursive(f, os.path.dirname(f), depth + 1)
            else:
                all_files.append(f)
    return all_files


def main():
    ap = argparse.ArgumentParser(description="递归解压询价 zip（中文文件名安全）")
    ap.add_argument("path", help="zip 文件或含 zip 的文件夹")
    ap.add_argument("-o", "--out", help="输出目录（默认 <输入同级>/_unzip）")
    args = ap.parse_args()

    base = args.path.rstrip("/")
    out = args.out or os.path.join(
        os.path.dirname(base) if os.path.isfile(base) else base, "_unzip")

    files = extract_recursive(base, out)
    print(f"\n共解出 {len(files)} 个非压缩文件，输出目录：{out}")
    for f in sorted(files):
        print("  ", os.path.relpath(f, out))


if __name__ == "__main__":
    main()
