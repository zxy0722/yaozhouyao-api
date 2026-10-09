# -*- coding: utf-8 -*-
"""
使用 Python 标准库（zipfile + xml）生成 .xlsx 文件，无需任何外部依赖。
输出: 耀州窑青釉刻花瓶_数据集.xlsx，包含 3 张工作表。
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "耀州窑青釉刻花瓶_数据集.xlsx")


# ===================================================================
# 工作表数据
# 每张表: (sheet_name, title_subtitle, [(序号, 数据类别, 字段, 值/描述, 来源), ...])
# ===================================================================

SHEET1 = (
    "表1·文物本体参数",
    "耀州窑青釉刻花瓶（北宋·新00143910）— 基本信息 / 尺寸参数 / 器型结构 / 釉色胎质 / 艺术效果",
    [
        (1,  "基本信息", "文物名称",     "耀州窑青釉刻花瓶",                                       "故宫博物院"),
        (2,  "基本信息", "文物编号",     "新00143910",                                            "故宫博物院"),
        (3,  "基本信息", "年代",         "北宋",                                                   "故宫博物院"),
        (4,  "基本信息", "窑口",         "耀州窑",                                                 "故宫博物院"),
        (5,  "基本信息", "收藏机构",     "故宫博物院",                                             "故宫博物院"),
        (6,  "尺寸参数", "通高",         "19.9厘米",                                               "故宫博物院"),
        (7,  "尺寸参数", "口径",         "6.9厘米",                                                "故宫博物院"),
        (8,  "尺寸参数", "足径",         "7.8厘米",                                                "故宫博物院"),
        (9,  "器型结构", "口部",         "平口出沿（小平口出沿）",                                 "故宫博物院"),
        (10, "器型结构", "颈部",         "短直颈",                                                 "故宫博物院"),
        (11, "器型结构", "肩部",         "丰肩",                                                   "故宫博物院"),
        (12, "器型结构", "腹部",         "腹以下渐收敛",                                           "故宫博物院"),
        (13, "器型结构", "足部",         "隐圈足（圈足）",                                         "故宫博物院"),
        (14, "器型结构", "肩部弦纹",     "弦纹3道",                                                "故宫博物院"),
        (15, "釉色胎质", "釉色",         "青釉（橄榄绿，青中带绿）",                               "陕西考古博物馆/耀州窑研究"),
        (16, "釉色胎质", "釉面质感",     "晶莹温润，玻璃质感强，半透明",                           "耀州窑址发掘报告"),
        (17, "釉色胎质", "釉层",         "匀净；较一般青瓷厚，透明度较弱；含密集微小气泡",         "Koh-antique 耀州窑研究"),
        (18, "釉色胎质", "胎色",         "灰白（北宋典型）",                                       "耀州窑址考古"),
        (19, "釉色胎质", "胎质",         "细腻致密，胎薄质坚，不吸水",                             "陕西省地方志办公室"),
        (20, "釉色胎质", "含铁量",       "较高（胎土富含氧化铁，强还原焰烧造后胎呈灰/黑灰色）",    "耀州窑青瓷研究"),
        (21, "釉色胎质", "圈足姜黄",     "圈足周围因胎含铁高＋煤窑氧化呈现姜黄色，为耀州青瓷标志",  "澎湃·观复博物馆"),
        (22, "艺术效果", "积釉效果",     "刻线处积釉色深，凸起处色较浅",                           "故宫博物院"),
        (23, "艺术效果", "图案效果",     "花纹清晰，有立体图案效果",                               "故宫博物院"),
    ],
)

SHEET2 = (
    "表2·工艺·纹饰·历史背景",
    "耀州窑青釉刻花瓶 — 工艺流程 / 装饰工艺 / 纹饰寓意 / 窑口背景 / 贡瓷窑神 / 地理交通 / 窑系影响",
    [
        # --- 工艺流程 / 装饰工艺 / 纹饰寓意 ---
        (24, "工艺流程", "总工序数",       "共17道：采料→精选→风化→配比→粑泥→陈腐→熟泥→揉泥→拉坯→修坯→釉料制备→施釉→装饰→窑具制作→装窑→烧窑→出窑", "耀州窑传统工艺研究"),
        (25, "工艺流程", "原料制备",       "瓷土采自黄堡镇附近二叠纪鸽青色及红黄色页岩，风化后使用；釉药为奥陶纪石灰岩中夹页岩，产自富平明月山",     "《同官县志·矿物志》"),
        (26, "工艺流程", "陈腐练泥",       "经采料、精选、风化、配比、粑泥、陈腐、熟泥、揉泥等工序",                                                "耀州窑传统工艺研究"),
        (27, "工艺流程", "成型方式",       "手工拉坯修坯（转轮就制，'方圆大小，皆中规矩'）",                                                       "北宋元丰七年《德应侯碑》"),
        (28, "装饰工艺", "装饰技法",       "刻花（斜刀深刻，'半刀泥'技法）",                                                                       "故宫博物院"),
        (29, "装饰工艺", "刻花工序",       "两步走：①斜刃直刀垂直深入行刀勾刻主轮廓；②下斜式行刀广削去胎，凸出花纹呈浮雕感",                       "耀州窑考古发掘报告"),
        (30, "装饰工艺", "细部处理",       "用多齿篦形工具在凸突主轮廓内勾划花瓣纹理与叶脉（划花配合）",                                            "禚振西《耀州窑青瓷的考古发现与鉴定》"),
        (31, "装饰工艺", "刻工特征",       "刀锋犀利，斜刀深刻",                                                                                    "故宫博物院"),
        (32, "装饰工艺", "工艺美誉",       "'巧如范金，精比琢玉'；'刀刀见泥'；'宋代青瓷刻花之冠'",                                                  "北宋《德应侯碑》/陕西考古博物馆"),
        (33, "装饰工艺", "主体纹饰",       "腹部满刻缠枝牡丹花",                                                                                    "故宫博物院"),
        (34, "装饰工艺", "近足纹饰",       "刻双层莲瓣纹",                                                                                          "故宫博物院"),
        (35, "装饰工艺", "牡丹纹特征",     "花繁而不乱，花冠丰满，俯仰结合",                                                                        "故宫博物院"),
        (36, "装饰工艺", "施釉方法",       "胎釉间施化妆土，通体施青釉，蘸釉或荡釉法",                                                              "陕西省地方志·耀州窑遗址"),
        (37, "工艺流程", "装烧工艺",       "匣钵内单件仰烧；或先素烧、上釉后再装烧的二次烧成法",                                                    "耀州窑址考古发掘"),
        (38, "工艺流程", "窑炉类型",       "馒头窑（马蹄形），耐火砖砌筑，由火膛、窑门、窑室、烟囱四部分组成",                                      "耀州窑遗址发掘"),
        (39, "工艺流程", "窑炉尺寸",       "窑室宽约2.16米、长约3.36米，每窑可装约2590个匣钵",                                                      "陕西省地方志办公室"),
        (40, "工艺流程", "燃料",           "煤（北宋始以煤代柴，提升烧成温度）",                                                                    "耀州窑考古"),
        (41, "工艺流程", "烧成温度",       "1280℃–1300℃",                                                                                          "耀州窑址考古报告"),
        (42, "工艺流程", "烧成气氛",       "由氧化焰转还原焰，烧出青绿光泽",                                                                        "耀州窑遗址发掘"),
        (43, "工艺流程", "成品率",         "约70%（无烟熏、烧生、开裂、变形）",                                                                     "陕西省地方志·耀州窑遗址"),
        (44, "纹饰寓意", "牡丹富贵",       "宋代牡丹称'富贵花'，象征繁荣昌盛、富贵吉祥",                                                            "《生活美学视域下宋代耀州窑陶瓷审美研究》"),
        (45, "纹饰寓意", "唐风延续",       "唐都长安提倡种植牡丹，宋延续此风，达官贵人与百姓皆喜观赏",                                              "同上"),
        (46, "纹饰寓意", "缠枝寓意",       "缠枝构图象征生生不息、子孙绵延、福禄绵长",                                                              "耀州窑纹饰研究"),
        (47, "纹饰寓意", "莲瓣佛教遗韵",   "双层莲瓣纹源于南北朝佛教纹样，象征清净高洁",                                                            "故宫博物院陶瓷研究"),
        (48, "纹饰寓意", "弦纹作用",       "肩部3道弦纹分隔颈、肩、腹三段，起界格与韵律作用",                                                       "耀州窑器型研究"),
        (49, "纹饰寓意", "因器施图",       "'纹饰和器物造型相统一'，瓶类结合腹部凸起处主体构图",                                                    "故宫博物院《陶瓷鉴赏》"),
        (50, "纹饰寓意", "题材类别",       "北宋耀州窑常见牡丹、莲、菊、梅、卷草等，牡丹运用最广",                                                  "《宋代耀州青瓷装饰纹样研究》"),
        (51, "纹饰寓意", "龙凤题材",       "龙凤纹为宫廷瓷器专用题材，民间不得使用",                                                                "陈炉古镇陶瓷文化介绍"),
        # --- 窑口背景 / 贡瓷窑神 / 地理交通 / 窑系影响 ---
        (52, "窑口背景", "创烧年代",     "唐代（黄堡窑场，初名'黄堡窑'）",                                                                                       "陕西省地方志·耀州窑遗址"),
        (53, "窑口背景", "鼎盛年代",     "北宋（960–1127），尤以北宋中期刻花技法成熟",                                                                            "耀州窑址考古"),
        (54, "窑口背景", "更名缘由",     "宋代黄堡划归耀州同官县治，窑以地名，故名'耀州窑'",                                                                     "《耀州窑址》百科"),
        (55, "窑口背景", "窑址位置",     "今陕西铜川黄堡镇一带，宋代属耀州",                                                                                     "故宫博物院"),
        (56, "窑口背景", "窑场规模",     "黄堡镇沿漆水河两岸密集布陈，史称'十里陶坊'；窑址范围长达5公里",                                                         "《耀州窑〈缠枝纹瓶〉》"),
        (57, "窑口背景", "窑场分布",     "中心窑场黄堡镇；另有立地坡、上店村、陈炉镇、玉华村等，绵延百里",                                                       "陕西省地方志"),
        (58, "贡瓷历史", "贡瓷烧造",     "宋神宗元丰（1078–1085）至徽宗崇宁（1102–1106）年间，曾为朝廷烧制贡瓷",                                                "新华网/《宋史·地理志》"),
        (59, "窑神封号", "德应侯",       "熙宁年间（1068–1077）耀州太守阎充国奏请，宋神宗封黄堡山神为'德应侯'——中国唯一皇封窑神",                              "北宋元丰七年《德应侯碑》"),
        (60, "窑神碑",   "《德应侯碑》", "立于元丰七年（1084）九月十八日，张隆撰书并题额；1974年移藏西安碑林博物馆，国家一级文物",                              "西安碑林博物馆"),
        (61, "碑文赞语", "范金琢玉",     "'巧如范金，精比琢玉……击其声铿铿如也，视其色温温如也'",                                                                "《德应侯碑》碑文"),
        (62, "地理交通", "京畿近镇",     "铜川唐时为长安京畿近镇，距长安约100余公里，近丝绸之路起点",                                                            "《耀州窑的地域文化》"),
        (63, "窑系影响", "耀州窑系",     "影响河南宜阳、宝丰、新安城关，广东广州西村，广西永福，内乡大窑店等窑场",                                              "《耀州窑址》百科"),
        (64, "外销证据", "海外出土",     "非洲一些国家出土耀州瓷文物，为通过古丝绸之路外销之明证",                                                              "新华网 2018-03-01"),
        (65, "鼎盛背景", "北宋中期断代", "北宋中期（仁宗至神宗，1023–1085）为刻花全盛期，器类大增，胎薄釉匀，纹饰满布器内外",                                   "耀州窑址考古报告"),
    ],
)

SHEET3 = (
    "表3·对比·鉴定·价值",
    "耀州窑青釉刻花瓶 — 同期对比 / 稀缺性 / 同类存世 / 鉴定特征 / 艺术评价 / 文化价值 / 文博地位 / 研究价值 / 传承现状",
    [
        (66, "同期对比", "与定窑对比",     "定窑'双刀一次完成'；耀州窑'单刀两步走'，斜削面比定窑宽",                                "人民网《古陶瓷装饰手法四大类》"),
        (67, "同期对比", "与越窑关系",     "五代始仿越窑青瓷与秘色瓷，划花风格近似，越窑对耀州窑发生过影响",                         "曾肃良《宋代耀州青瓷的纹饰风格与意义》"),
        (68, "同期对比", "与汝窑关系",     "五代耀州天青釉瓷先于汝窑，对其工艺风格产生深远影响",                                    "王小蒙《耀州窑天青釉瓷研究》"),
        (69, "稀缺性",   "器类稀缺",       "耀州窑刻花青瓷以盘、碗器皿居多，瓶类较少；本瓶为稀缺器型",                              "故宫博物院藏本说明"),
        (70, "同类存世", "缠枝牡丹梅瓶",   "故宫博物院另藏'耀州窑青釉刻缠枝牡丹纹梅瓶'（宋）",                                      "故宫博物院数字文物库"),
        (71, "同类存世", "牡丹纹执壶",     "1989年陕西耀县出土'耀州窑青釉刻牡丹花卉纹执壶'（北宋），藏陕西耀州窑博物馆",            "澎湃·学术"),
        (72, "同类存世", "套盒",           "2009年西安蓝田吕氏家族墓地出土'耀州窑青釉刻牡丹纹套盒'（北宋），藏陕西省考古研究院",    "澎湃·学术"),
        (73, "鉴定特征", "真品刀痕",       "斜刃/平刃直刀两步走，刀锋犀利圆活；北宋晚期出现圆圜刃刀，刻纹较浅、动感圆活",                "禚振西《后仿耀州窑瓷器的装饰鉴定》"),
        (74, "鉴定特征", "后仿破绽",       "现代仿品多用90°拐角刀一次刻成，线条生硬，缺犀利圆活之美",                                  "同上"),
        (75, "鉴定特征", "圈足姜黄",       "圈足周围姜黄色为耀州青瓷标志性特征，可作为断代鉴识依据",                                    "澎湃·观复博物馆"),
        (76, "艺术评价", "器型评价",       "小口短颈衬出瓶身雍容饱满",                                                                  "故宫博物院"),
        (77, "艺术评价", "北方青瓷代表",   "耀州窑被誉为'北方青瓷之首/代表'",                                                            "《宋代耀州青瓷装饰纹样研究》"),
        (78, "艺术评价", "刻花之冠",       "北宋耀州窑刻花被誉为'宋代青瓷刻花之冠'",                                                    "西安晚报·陕西考古博物馆"),
        (79, "艺术评价", "风格定位",       "构图饱满，刀法灵锐，走线洒脱，颇具北方地域色彩",                                            "观复博物馆"),
        (80, "艺术评价", "审美取向",       "体现宋代'尚玉'审美——青釉色与玉温润相似，受国人崇玉影响",                                    "澎湃·橄榄青的韵味"),
        (81, "艺术评价", "总体定位",       "为耀州窑瓷器中的精品",                                                                       "故宫博物院"),
        (82, "文化价值", "制器理念",       "'物以致用'+'制器尚象'，造型与功能统一，自然元素灵活运于器型",                                "《生活美学视域下宋代耀州窑陶瓷审美研究》"),
        (83, "文化价值", "生活美学",       "'艺术生活化'与'生活艺术化'，雅俗共赏、兼容并蓄",                                            "同上"),
        (84, "文化价值", "丝绸之路器物",   "经丝绸之路远销海外，为中华文化传播载体",                                                    "新华网"),
        (85, "文博地位", "收藏机构等级",   "故宫博物院（国家一级博物馆）",                                                              "原参数"),
        (86, "研究价值", "断代意义",       "耀州窑址发掘是中国首次对古瓷窑址的大面积发掘，对北方青瓷断代有重要意义",                    "陕西省地方志"),
        (87, "传承现状", "炉火延续",       "黄堡窑场明中期断烧，陈炉窑场延续至今近1400年，为'活着的耀州窑'",                            "陈炉古镇陶瓷文化介绍"),
    ],
)

SHEETS = [SHEET1, SHEET2, SHEET3]


# ===================================================================
# xlsx 生成（纯标准库实现）
# ===================================================================
COL_WIDTHS = [6, 12, 16, 70, 28]  # 序号/数据类别/字段/值/来源


def col_letter(idx):
    """1-based index -> A, B, ..., Z, AA, ..."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def xml_attr_escape(s):
    return escape(s).replace('"', "&quot;").replace("'", "&apos;")


def build_content_types():
    overrides = []
    overrides.append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    overrides.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    overrides.append('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    overrides.append('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
    for i in range(1, len(SHEETS) + 1):
        overrides.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    overrides.append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    overrides.append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'{"".join(overrides)}'
        '</Types>'
    )


def build_root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml" />'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml" />'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml" />'
        '</Relationships>'
    )


def build_workbook_xml():
    """xl/workbook.xml - r:id 必须与 workbook.xml.rels 中 worksheet 关系的 Id 一致。"""
    sheets_xml = []
    for i, (name, _, _) in enumerate(SHEETS, start=1):
        rid = f"rId{i+3}"  # 与 build_workbook_rels 中工作表的 rId 保持一致
        sheets_xml.append(
            f'<sheet name="{xml_attr_escape(name)}" sheetId="{i}" state="visible" r:id="{rid}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/x15/15/main">'
        '<workbookPr/>'
        f'<sheets>{"".join(sheets_xml)}</sheets>'
        '<definedNames/>'
        '<calcPr calcId="0"/>'
        '</workbook>'
    )


def build_theme_xml():
    """xl/theme/theme1.xml - 最小化主题，WPS/Excel 期待此部件存在。"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'name="Office 主题">'
        '<a:themeElements>'
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
        '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
        '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
        '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
        '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
        '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
        '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Office">'
        '<a:majorFont>'
        '<a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/>'
        '</a:majorFont>'
        '<a:minorFont>'
        '<a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>'
        '</a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
        '</a:gsLst></a:gradFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '<a:objectDefaults/>'
        '</a:theme>'
    )


def build_workbook_rels():
    """xl/_rels/workbook.xml.rels - 工作簿到工作表的关系。
    Type 必须是 worksheet（不是 officeDocument，那是 .rels->workbook.xml 的类型）。
    需补充 styles/sharedStrings/theme 的关系，否则 WPS 无法定位。"""
    rels = []
    # styles
    rels.append(
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml" />'
    )
    # sharedStrings
    rels.append(
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml" />'
    )
    # theme (WPS/Excel 期待此部件存在，否则可能空白)
    rels.append(
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml" />'
    )
    # worksheets
    for i in range(1, len(SHEETS) + 1):
        rid = f"rId{i+3}"
        rels.append(
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml" />'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(rels)}</Relationships>'
    )


def build_core_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>耀州窑青釉刻花瓶 数据集</dc:title>'
        '<dc:creator>Trae</dc:creator>'
        '<cp:lastModifiedBy>Trae</cp:lastModifiedBy>'
        '</cp:coreProperties>'
    )


def build_app_xml():
    titles = "".join(f'<vt:lpstr>{xml_attr_escape(n)}</vt:lpstr>' for n, _, _ in SHEETS)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Trae</Application>'
        f'<TitlesOfParts><vt:vector size="{len(SHEETS)}" valueType="lpstr">{titles}</vt:vector></TitlesOfParts>'
        '</Properties>'
    )


def build_styles_xml():
    """样式: 0=默认; 1=标题(粗体居中+深蓝填充); 2=表头(粗体居中+中蓝填充); 3=数据(自动换行+顶对齐); 4=序号居中浅蓝"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="5">'
        '<font><sz val="11"/><name val="等线"/></font>'
        '<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="等线"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/></font>'
        '<font><sz val="11"/><name val="等线"/></font>'
        '<font><b/><sz val="11"/><name val="等线"/></font>'
        '</fonts>'
        '<fills count="5">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>'
        '</fills>'
        '<borders count="2">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border>'
        '<left style="thin"><color rgb="FFBFBFBF"/></left>'
        '<right style="thin"><color rgb="FFBFBFBF"/></right>'
        '<top style="thin"><color rgb="FFBFBFBF"/></top>'
        '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom>'
        '</border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="5">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def build_shared_strings_xml(pool):
    items = []
    for s in pool:
        ph = escape(s).replace("\n", "&#10;")
        items.append(f'<si><t xml:space="preserve">{ph}</t></si>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(pool)}" uniqueCount="{len(pool)}">'
        f'{"".join(items)}</sst>'
    )


def build_sheet_xml(sheet_idx, idx_map):
    name, subtitle, data_rows = SHEETS[sheet_idx]
    n_cols = 5
    last_col = col_letter(n_cols)
    n_rows = 1 + 1 + len(data_rows)

    cols_xml = "".join(
        f'<col min="{i+1}" max="{i+1}" width="{COL_WIDTHS[i]}" customWidth="1"/>'
        for i in range(n_cols)
    )

    row_xmls = []

    # Row 1: 标题(合并 A1:E1)
    title_text = f"{name}：{subtitle}"
    title_si = idx_map[title_text]
    row_xmls.append(
        f'<row r="1" spans="1:{n_cols}" ht="36" customHeight="1">'
        f'<c r="A1" t="s" s="1"><v>{title_si}</v></c>'
        f'<c r="B1" s="1"/>'
        f'<c r="C1" s="1"/>'
        f'<c r="D1" s="1"/>'
        f'<c r="E1" s="1"/>'
        f'</row>'
    )

    # Row 2: 表头
    headers = ["序号", "数据类别", "字段", "值/描述", "来源"]
    cells = ""
    for ci, h in enumerate(headers, start=1):
        ref = f"{col_letter(ci)}2"
        si = idx_map[h]
        cells += f'<c r="{ref}" t="s" s="2"><v>{si}</v></c>'
    row_xmls.append(f'<row r="2" spans="1:{n_cols}" ht="22" customHeight="1">{cells}</row>')

    # Row 3+: 数据
    for ri, row in enumerate(data_rows, start=3):
        cells = ""
        for ci, val in enumerate(row, start=1):
            ref = f"{col_letter(ci)}{ri}"
            s = "4" if ci == 1 else "3"
            si = idx_map[str(val)]
            cells += f'<c r="{ref}" t="s" s="{s}"><v>{si}</v></c>'
        row_xmls.append(f'<row r="{ri}" spans="1:{n_cols}" ht="48" customHeight="1">{cells}</row>')

    merges = f'<mergeCell ref="A1:{last_col}1"/>'

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        f'<dimension ref="A1:{last_col}{n_rows}"/>'
        f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/>'
        f'<cols>{cols_xml}</cols>'
        f'<sheetData>{"".join(row_xmls)}</sheetData>'
        f'<mergeCells count="1">{merges}</mergeCells>'
        '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
        '<pageSetup orientation="portrait" paperSize="9"/>'
        '</worksheet>'
    )


def main():
    # 构造共享字符串池
    pool = []
    idx_map = {}

    def add(v):
        if v not in idx_map:
            idx_map[v] = len(pool)
            pool.append(v)
        return idx_map[v]

    # 标题字符串(每表独立)
    for n, s, _ in SHEETS:
        add(f"{n}：{s}")
    # 表头
    for h in ["序号", "数据类别", "字段", "值/描述", "来源"]:
        add(h)
    # 数据
    for _, _, rows in SHEETS:
        for r in rows:
            for v in r:
                add(str(v))

    # 写入 zip
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", build_content_types())
        z.writestr("_rels/.rels", build_root_rels())
        z.writestr("xl/workbook.xml", build_workbook_xml())
        z.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels())
        z.writestr("xl/sharedStrings.xml", build_shared_strings_xml(pool))
        z.writestr("xl/styles.xml", build_styles_xml())
        z.writestr("xl/theme/theme1.xml", build_theme_xml())
        for i, _ in enumerate(SHEETS):
            z.writestr(f"xl/worksheets/sheet{i+1}.xml", build_sheet_xml(i, idx_map))
        z.writestr("docProps/core.xml", build_core_xml())
        z.writestr("docProps/app.xml", build_app_xml())

    print("OK ->", OUT)
    print("Sheets:")
    for n, s, rows in SHEETS:
        print(f"  - {n} ({len(rows)} 行)")


if __name__ == "__main__":
    main()
