# PBOM 中文版 · L463 6座出入口盖板（注塑件）

> 原文件：`L463-010501- Injection moulded PANEL PBOM Rev1.pptx`（PBOM＝Production Bill of Materials，**生产物料清单**）
> 说明：只翻译技术相关文字，供技术员看懂；**零件号/材料牌号/数模编号/文件名保持英文原样**，后面加中文释义。
> 生成：2026-08-19

---

## 第 1 页 · 封面
- **L463-010501 注塑出入口盖板 PBOM（生产物料清单）**
- 日期：2026-07-15

## 第 2 页 · 总体说明（General Notes）
标题：L463-010501 **TPO 包覆出入口盖板 PBOM**
- 各零件的**单车用量（take rate）和产量**见 eSOW 成本包（ESOW cost pack）文件。
- 零件几何形状**随数模（CAD）逐步成熟可能会变更**（还没冻结）。
- 图上标的尺寸**只是外形包围盒（最大外廓 bounding box）尺寸**，**一切以数模（CAD）为准**。

## 第 3 页 · 爆炸图：6座出入口盖板（总成）
- **爆炸视图（Exploded View）**：6 SEAT ACCESS PANEL（6座出入口盖板）
- 数模编号（CAD REFERENCE）：`EI_ Z7B2-13A994-AAW`

**零件清单（BOM）：**

| 序号 | 零件名称 | 数量 | 备注（工艺）|
|---|---|---|---|
| 1 | 6 SEAT ACCESS PANEL 6座出入口盖板（本体）| 1 | **IM Process＝注塑成型工序** |
| 2 | ITW MAT CLIP（ITW 品牌垫用卡扣）| 1 | **Clip Assy Process＝卡扣装配工序** |
| 3 | W780230-S CLIP STANDARD（W780230-S 标准卡扣）| 8 | 卡扣装配工序 |

**材料信息：**
- **拟用材料（Proposed Material）：ABS MAGNUM 3325**（一种 ABS 塑料牌号）
- **颜色（Colour）：Shadow**（暗影色）
- **皮纹（Grain）：TBC**（To Be Confirmed＝**待客户确认**）

## 第 4 页 · 爆炸图：6座出入口盖板 — ITW 卡扣
- 爆炸视图：6座出入口盖板 —— **ITW 卡扣**（展示 ITW 卡扣的装配关系）

## 第 5 页 · 爆炸图：6座出入口盖板 — ITW 卡扣
- 同上，**ITW 卡扣**装配细节（另一角度/局部）

## 第 6 页 · 爆炸图：6座出入口盖板 — 注塑
- 爆炸视图：6座出入口盖板 —— **INJECTION MOULDING（注塑成型）**
- 数模编号（CAD REFERENCE）：`EI_ Z7B2-13A994-AAW`
- **模穴假设（CAVITY ASSUMPTION）：1 CAV TOOL＝一模一穴（一模一出）模具**
- ＊图上尺寸为外形包围盒（最大外廓）尺寸

## 第 7 页 · 爆炸图：6座出入口盖板（总装）
- 爆炸视图：6座出入口盖板
- 数模编号（CAD REFERENCE）：`EI_ Z7B2-13A994-AAW`
- ＊图上尺寸为外形包围盒尺寸
- **模穴假设：一模一穴（1 CAV TOOL）**
- **1x MAT CLIP ASSY**＝1 个垫用卡扣组件（ITW）
- **8x METAL CLIP ASSY**＝8 个金属卡扣组件（W780230-S）

---

## 附：页面里「零件数据表」图片中的技术数据（图中文字，一并翻译）

> 第 6/7 页的数据表是**图片**（不在可复制文字里），内容对技术员很关键，一并译出：

| 英文字段 | 中文 | 值 |
|---|---|---|
| Part name | 零件名称 | 6 SEAT ACCESS PANEL 6座出入口盖板 |
| Part number | 零件号 | Z7B2-13A994-A-INS-01 |
| Number LH+RH / car | 单车左右件合计数量 | 1 |
| Mirrored part (left/right) | 是否左右对称件 | 否（no）|
| Platform | 车型平台 | L463（路虎 Defender）|
| Length (X) | 长（X向）| 490 mm（measured 实测）|
| Width (Y) | 宽（Y向）| 329 mm（实测）|
| Height (Z) | 高（Z向）| 52 mm（实测）|
| Weight | 重量 | 0.450 kg（calculated 计算值）|
| Surface Area | 表面积 | 0.145~0.147 m²（计算值）|
| Material | 材料 | ABS Magnum 3325（另一版数据表写 Borealis Borcycle MG2503SY）|
| Manufacturing techn. | 制造工艺 | Injection Moulding 注塑成型 |
| Material thickn. | 料厚（壁厚）| 2.5 mm（avg 平均）|
| Surface | 表面处理 | grained 皮纹（另一版写 raw 素面/未处理）|
| Color / type of grain | 颜色 / 皮纹类型 | Grain and Pattern TBC by JLR（皮纹与纹路待路虎确认）|

---

### 术语速查（这份文件里的英文缩写）
- **IM＝Injection Moulding**：注塑成型
- **PBOM**：生产物料清单
- **CAD / 数模**：三维数字模型，几何形状以它为准
- **Bounding box**：外形包围盒 = 零件最大外廓的长×宽×高
- **CAV / Cavity**：模穴，「1 CAV」= 一模一出
- **Clip Assy**：卡扣装配
- **TBC（To Be Confirmed）**：待确认
- **grained**：皮纹面（有纹理的外观面）｜**raw**：素面/未做纹理｜**wrapped**：包覆（贴皮）｜**painted**：喷漆
- **TPO**：一种可包覆用的软质料（TPO 包覆版盖板表面贴 TPO 皮）
