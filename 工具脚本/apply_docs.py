#!/usr/bin/env python3
# 双语标注: 2个docx + F16/EJ/KM 三个pptx。原文保留,后/下加中文。材料牌号/件号/代码不译。
import sys, os
sys.path.insert(0, "/Users/chensiyuan5/Desktop/Github/shenyuan_work/工具脚本")
import bilingual as B

BASE = "/Users/chensiyuan5/Desktop/Github/shenyuan_work/注塑项目/金榕/未报价 镜面覆盖塑料件 2"
KM = BASE + "/｣ｨKM -Mirror Protector -2026｣ｩ---Technical Solution"
EJ = BASE + "/EJ Door Bags and Straps---Technical Solution"

# ---------- Mirror protector - FINAL.docx ----------
mirror_final = {
    "Mirror protector:": "镜面保护罩:",
    "Function: Prevent the Glass mirror theft.": "功能:防止玻璃后视镜被盗。",
    "Material: Acrylonitrile Styrene Acrylate (ASA BLACK) / ABS": "材料:ASA BLACK / ABS(材料牌号保留原文,采购用)",
    "Texture: COD.001": "纹理:COD.001(纹理代码,注意PPT写的是COD.003,需向客户确认)",
    "Adhesive: Epoxy (EP-20/H-97M)": "胶水:Epoxy (EP-20/H-97M)(牌号保留原文)",
    "Fitment:": "装配要求:",
    "Part should have at least 2mm of gap against all mirror adjustment positions.":
        "零件与后视镜所有调节位置之间至少留 2mm 间隙。",
    "Should fit along all edges. Should not interfere with any mirror functionality.":
        "应沿所有边缘贴合;不得干涉后视镜的任何功能。",
    "Vehicle: B10 & C10/C16": "车型:B10 & C10/C16",
    "P/N: 82219786AA": "件号:82219786AA",
    "Weight: 0.21 kg": "重量:0.21 kg",
    "Pieces per set: 2# - 1 - LH & 1 - RH": "每套件数:2 件 — 左(LH)1 件 + 右(RH)1 件",
    "Process: Injection Molding": "工艺:注塑成型",
}

# ---------- Statement of Work and ED.docx ----------
sow = {
    "Statement of Work and ED&D": "工作说明书与工程设计开发(ED&D)",
    "The chosen supplier will be responsible for design, development, PPAP and supply of the kit. The part needs to be developed based on the Artwork provided by our product design office. The supplier needs to produce sample parts and to be installed on vehicle provided by us for on vehicle approval by us. This may take few iterations as sometimes it may take few adjustments due to complex surface geometries.":
        "选定的供应商负责整套件的设计、开发、PPAP 及供货。零件须基于我方产品设计室提供的外观数据(Artwork)开发。供应商须制作样件,装到我方提供的车辆上做整车认可。由于曲面复杂,可能需要多次迭代微调。",
    "The supplier will also be responsible for creating 3D data, 2D drawings & I-sheet for the service parts. The film material must meet our specifications and supplier will be responsible to provide surrogate data/panels or must carry new testing if any data is not available.":
        "供应商还负责服务件的 3D 数据、2D 图纸和 I-sheet(说明书)的创建。薄膜材料(film material)须满足我方规范;若缺少相关数据,供应商负责提供替代数据/样板或重新做测试。",
}

# ---------- F16 pptx ----------
f16 = {
    "DT – F16 Ram Sports Truck End Cap (SPAT) Development – Mopar": "DT-F16 Ram 运动皮卡 端盖(气坝SPAT)开发-Mopar",
    "BY JINRONG": "由金榕发包",
    "PRODUCT DEFINITION AND SCOPE OF DESIGN & SUPPLY": "产品定义与设计供货范围",
    "ITEM": "项目", "DETAILS": "内容",
    "Part Name": "零件名称",
    "Spat Kit Right & Left - Primed finish, ready for paint.": "气坝套件 左+右 — primed底涂,可直接喷漆",
    "New Part Number": "新件号",
    "Vehicle & MYI-MYO": "车型与年款(MYI-MYO)",
    "DT Sports Truck with front Bumper (Sales Code MCE)": "DT 运动皮卡 带前保险杠(销售代码 MCE)",
    "Stage 1 (5.7L) & Stage 2A (6.4L) base": "Stage 1(5.7L)与 Stage 2A(6.4L)基础版",
    "MY 2027-2029": "年款 2027-2029",
    "Vehicle Application": "应用", "Appearance & Performance": "外观件 & 功能件",
    "Installation Location": "安装位置",
    "Front Bumper, right and left ends. At dealership – post sale": "前保险杠左右两端;经销商端售后加装",
    "Front Bumper, right and left ends.": "前保险杠左右两端",
    "Material Grade & Color Code & Finish type & Texture & Color tolerance": "材料牌号 & 颜色代码 & 表面处理 & 纹理 & 色差公差",
    "Annual Volume": "年用量",
    "Spat (light blue color)": "气坝(浅蓝色)",
    "Retail service part definition": "零售服务件定义",
    "Spat Kit Right & Left – Painted, Body Color": "气坝套件 左+右 — 已喷漆,车身同色",
    "Part installed at UTM Custom Shop, Saltillo Mexico": "在墨西哥 Saltillo 的 UTM 定制车间装配",
    "Material, Color, finish": "材料、颜色、表面处理",
    "Same as 82219733AA but painted body color as per above": "同 82219733AA,但按上表喷成车身色",
    "Custom shop parts definition": "定制车间件定义",
    "PN": "件号", "Description": "描述", "Annual volume": "年用量", "Material": "材料", "QTY": "数量",
    "ZR6 - Molten Red Pearl Coat": "ZR6 - 熔岩红珠光漆",
    "ZG9 - Serrano Green Metallic": "ZG9 - 塞拉诺绿金属漆",
    "KXJ - Diamond Black Crystal": "KXJ - 钻石黑水晶",
    "PDN - Ceramic Grey Clear Coat": "PDN - 陶瓷灰清漆",
    "GW7 - Bright White Clear Coat": "GW7 - 亮白清漆",
    "EYB - Detonator Yellow": "EYB - 引爆黄",
    "Scope of Work": "工作范围",
    "The scope of work will be design, development, manufacture, PPAP and supply of the parts.":
        "工作范围:零件的设计、开发、制造、PPAP 与供货。",
    "The expectation is to deliver parts to:": "交付预期:",
    "Retail: 82219733AA (in prime) individually packaged to Mopar Depot.":
        "零售件:82219733AA(primed底涂),单件包装发往 Mopar 仓库。",
    "Custom Shop: Six part numbers, in body color, to Utilimaster custom shop in Saltillo Mexico":
        "定制车间件:6 个件号,车身色,发往墨西哥 Saltillo 的 Utilimaster 定制车间。",
    "The intent is to use same attachment strategy and the existing hardware herever possible from the current existing similar production parts.":
        "尽量沿用现有相似量产件的连接方式与现有五金件。",
    "Jinrong will provide direct buy letters to purchase the existing or any current production parts from the existing production suppliers.":
        "金榕将提供定向采购(direct buy)函,供向现有量产供应商采购现有件。",
    "The build objectives for the Gaps and finish will be same as the similar current production parts for this vehicle.":
        "间隙与表面质量目标与本车现有相似量产件一致。",
    "The A surface CAD data will be provided after the award of business and supplier will be responsible for development of B-surface and attachment points.":
        "A 面 CAD 数据在定点后提供;供应商负责 B 面及连接点的开发。",
    "The selected supplier will be responsible for creation of CAD for the new part and 2D drawings.":
        "供应商负责新件的 CAD 与 2D 图纸创建。",
    "A tech review can be requested if there are any questions or concerns prior to submission of quotes.":
        "报价前如有疑问可申请技术评审(tech review)。",
    "All component level testing will be supplier responsibility, and the system and vehicle level testing will be carried out by the Jinrong or Supplier.":
        "零件级测试由供应商负责;系统级与整车级测试由金榕或供应商进行。",
    "Performance/test requirements are available in the DVP&R attached with the SP.":
        "性能/测试要求见随 SP 附的 DVP&R。",
    "The chosen supplier will be responsible for supply of parts for the marketing and testing. Protect for 3 set of parts.":
        "供应商需供应市场宣传与测试用件,备 3 套。",
    "The DVP&R states the standard requirements for new parts, no testing will be required for the directed buy production components. Effort should be made to use products which are proven, and test data is available to minimize the new testing wherever possible.":
        "DVP&R 规定新件标准要求;定向采购的量产件不需重新测试。尽量选用已验证、有测试数据的产品以减少新测试。",
    "PART SUMMARY": "零件小结", "OTHER": "其它",
    "Final data must be submitted to Mopar and reviewed/approved prior to kick-off.":
        "最终数据须提交 Mopar 评审/批准后方可启动(kick-off)。",
    "Supplier is design responsible and must provide CAD models for this PN.":
        "供应商承担设计责任,须提供该件 CAD 模型。",
    "Test costs must be submitted and itemized based on DVP, as well as additional ED&D (Engineering Design & Development) costs for engineering or design services":
        "测试费须按 DVP 逐项列出,另含工程设计开发(ED&D)费用。",
    "GD&T documentation and quote assumptions must follow guidelines defined in PS-9611 Geometric Dimensioning and Tolerancing Practice.":
        "GD&T(几何尺寸与公差)文档及报价假设须遵循 PS-9611 规范。",
    "Supplier is responsible for the development of installation instructions.":
        "供应商负责编写安装说明书。",
    "Supplier is Responsible for Development of I-Sheet. Kit will contain QR code sheet provided by Jinrong, which links to the I-sheet document electronically.":
        "供应商负责开发 I-Sheet;套件含金榕提供的二维码卡,链接到电子版 I-sheet。",
    "VALIDATION:": "验证:",
    "Must satisfy requirements detailed in the DVP&R and the following Jinrong standards.":
        "须满足 DVP&R 及以下金榕标准的要求。",
    "If supplier intends to use surrogate data for any of the DVP line items, it must be noted at time of quote submittal, reviewed with engineering, and should not be more than 3 years old, unless a deviation is provided by Mopar Product Development.":
        "若某 DVP 项拟用替代数据,报价时须注明并与工程评审,且数据不超过 3 年,除非 Mopar 产品开发出具偏差许可。",
    "Part Volume - Sample and SoP:": "样件与量产数量:",
    "Engineering Fit-ups: 3 samples (primed)": "工程试配:3 件(primed底涂)",
    "Testing/Validation as required per attached DVP": "按所附 DVP 做测试/验证",
    "Supplier Quality: 3 kits - Marketing: 1": "供应商质量:3 套;市场:1 套",
    "PER samples: 2 body colors": "PER 样件:2 种车身色",
    "Timing detail:": "时间计划明细:",
    "Supplier must provide a detailed timeline and should capture specifics on all the activities related to this development, including but not limited to PO's, Design, Testing, Tooling, DV testing, system level PV testing, Packaging, PPAP,":
        "供应商须提供详细时间表,覆盖全部开发活动,包括但不限于:PO、设计、测试、模具、DV测试、系统级PV测试、包装、PPAP,",
    "Logistics, and Shipment to Mopar warehouse Timeline should capture responsible person/group for each task and should capture dependencies.":
        "物流、发运至 Mopar 仓库。时间表须标明每项任务的责任人/组及依赖关系。",
    "Timing should also include any contingencies for additional factors (holidays, shutdowns, sea shipments, second  development etc.) to meet the vehicle timing requirements listed above":
        "时间表还须考虑节假日、停产、海运、二次开发等余量,以满足上述整车时间要求。",
    "Timeline sent should be accompanied with a document listing all assumptions that won't into defining the durations for the tasks documented. Any anticipated risks to the proposed timeline should be highlighted along with proposals to address the same":
        "时间表须附假设清单;并标出预期风险点及对策。",
    "A detailed process depicting location and responsible source for raw materials, tool design, tool build, part manufacture, part development, part packaging, and where shipped from should accompany the timeline.":
        "须附详细流程:原材料、模具设计/制造、零件制造/开发/包装的地点与责任来源及发运地。",
    "Timing and sample requirements": "时间与样件要求",
    "Left Side Part Shown": "图示为左侧件", "Spat": "气坝 Spat", "Bracket": "支架",
    "Mopar 27DT Front Aero Spat Presource Info05152026": "Mopar 27DT 前导流气坝 预采购信息 05152026",
    "Plan View": "俯视图", "Front View": "前视图", "Right Side View": "右视图",
    "Left Side View": "左视图", "Rear View": "后视图",
    "Spacer": "垫片", "4.2 mm ID": "内径 4.2mm", "15 to 25 mm OD": "外径 15~25mm",
    "5 mm thick": "厚 5mm", "Close Out": "封边", "Production splash shield": "量产挡泥板",
    "(reference only)": "(仅供参考)", "Fasteners": "紧固件",
    "M-4.8 U-Clip, End Cap to wheel liner": "M-4.8 U型卡扣,端盖固定到轮罩衬板",
}

# ---------- EJ pptx ----------
ej = {
    "KIT BOM": "套件物料清单(BOM)",
    "S.No": "序号", "Part name": "零件名称", "Qty": "数量", "Box dim.": "包装盒尺寸",
    "Buckle": "带扣", "Nylon Cords/Straps": "尼龙绳/绑带", "Fastener": "紧固件", "I-sheet": "说明书",
    "NA": "无",
    "Attachment method of current buckle & straps to Door for reference in below images,":
        "下图为现有带扣与绑带到车门的连接方式,仅供参考,",
    "( Dimensions are (in ‘mm’) approximate and for referential purposes only)":
        "(图中尺寸单位 mm,为近似值,仅供参考)",
}

# ---------- KM 纯图片 PPT:叠加中文文本框 ----------
km_overlay = {
    1: "镜面保护罩 KM(图2为产品定义):功能=防玻璃后视镜被盗;材料=ASA BLACK/ABS(原文不译);纹理=COD.003(注意FINAL.docx写COD.001);"
       "胶水=Epoxy;装配=与后视镜各调节位至少2mm间隙,沿边贴合不干涉功能;车型KM;颜色黑;件号82219786AA;重量TBD;每套2件(左+右);工艺=注塑成型。",
}

print("docx: Mirror FINAL ...")
B.annotate_docx(KM + "/Mirror protector - FINAL.docx", KM + "/Mirror protector - FINAL.docx", mirror_final)
print("docx: SoW ...")
B.annotate_docx(KM + "/Statement of Work and ED.docx", KM + "/Statement of Work and ED.docx", sow)
print("pptx: F16 ...")
B.annotate_pptx(BASE + "/F16 End Cap_Spat  tech sheet_5-18-2026 .pptx",
                BASE + "/F16 End Cap_Spat  tech sheet_5-18-2026 .pptx", f16)
print("pptx: EJ ...")
B.annotate_pptx(EJ + "/27EJ Door Bags-tech sheet.pptx", EJ + "/27EJ Door Bags-tech sheet.pptx", ej)
print("pptx: KM overlay ...")
B.annotate_pptx(KM + "/KM -Mirror Protector -2026-tech-sheet.pptx",
                KM + "/KM -Mirror Protector -2026-tech-sheet.pptx", {}, slide_overlays=km_overlay)
print("DONE")
