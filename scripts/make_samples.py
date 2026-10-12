"""生成一套模拟数据（虚构的“青石山”“白云岭”，数据均为编造），用来演示和测试。

生成内容：资料库/ 2 份资料、模板/ 空白模板、样例/ 正确样例、待审核/ 两份含错误的填报表、默认 审核规则.xlsx
用法：python make_samples.py
"""
import config
from build_rules import save_rules
from rules import DEFAULT_COLUMNS
from table import write_table

FILLER = "本节其余内容记述了相关历史沿革、民间传说及开发利用情况，与上述数据无直接关系，此处从略。" * 3

COUNTY_ANNALS = f"""青石县志（2018年版）
青石县地方志编纂委员会 编
2018年6月 出版

第一章 自然地理
第一节 山脉
青石山为县内最高峰，主峰海拔1203米，位于东经102°30′15″、北纬25°10′05″。
山体呈东北—西南走向，西坡陡峭、东坡平缓，山顶发育喀斯特峰丛。
{FILLER}
白云岭位于县境东南，主峰海拔876米，位于东经102°41′30″、北纬25°03′20″。
白云岭属低山丘陵，坡度平缓，多为红壤缓坡，山间有小块河谷盆地。
{FILLER}
"""

VEGETATION_REPORT = f"""青石县植被调查报告（2019年）
青石县林业和草原局
2019年11月

一、青石山
青石山海拔900米以下以常绿阔叶林为主，900米以上为针阔混交林，山顶有少量灌草丛。
{FILLER}
二、白云岭
白云岭植被以云南松林为主，山脚缓坡分布有较大面积的人工经济林（核桃、板栗）。
{FILLER}
"""

ITEMS = [  # 数据类型, 数据要求, 字符要求
    ("名称", "山体或山脉的规范名称", "纯汉字，2~10个字"),
    ("高度", "主峰海拔，单位：米，填整数", "纯数字"),
    ("地理坐标", "主峰位置，经度在前、纬度在后，度分秒或十进制度均可", "经纬度"),
    ("地形特点", "简要描述山体走向、坡度和地貌类型", "不超过100字"),
    ("覆盖植被类型", "列出主要植被类型，用顿号分隔", "不超过50字"),
]

ANNALS, VEG = "《青石县志（2018年版）》", "《青石县植被调查报告（2019年）》"

QINGSHI_OK = [  # 数据内容, 数据来源, 来源日期
    ("青石山", ANNALS, "2018-06"),
    (1203, ANNALS, "2018-06"),
    ("东经102°30′15″，北纬25°10′05″", ANNALS, "2018-06"),
    ("山体呈东北—西南走向，西坡陡峭、东坡平缓，山顶为喀斯特峰丛", ANNALS, "2018-06"),
    ("常绿阔叶林、针阔混交林、灌草丛", VEG, "2019-11"),
]

# 甲单位：青石山，每行埋一种典型错误（地理坐标一行正确，用十进制度写法测试换算）
UNIT_A = [
    ("青石山", ANNALS, "2019-06"),                                # 来源日期错：县志是2018年6月
    (120, ANNALS, "2018-06"),                                    # 数值错：原文是1203米
    ("东经102.504°，北纬25.168°", ANNALS, "2018-06"),             # 正确（度分秒换算后一致）
    ("山体呈南北走向，东坡陡峭、西坡平缓", ANNALS, "2018-06"),       # 走向和坡向都与原文相反
    ("常绿阔叶林、竹林", VEG, "2019-11"),                          # 部分支持：原文没有竹林
]

# 乙单位：一张表两座山。青石山全部正确；白云岭有错
UNIT_B = QINGSHI_OK + [
    ("白云岭", ANNALS, "2018-06"),                                # 正确
    (768, ANNALS, "2018-06"),                                    # 数字颠倒：应为876
    ("东经25°03′20″，北纬102°41′30″", ANNALS, "2018-06"),          # 经纬度写反
    ("低山丘陵，坡度平缓", ANNALS, "2018-06"),                      # 正确（概括原文）
    ("云南松林、人工经济林", ANNALS, "2018-06"),                    # 引错资料：县志没写植被，应引植被报告
]

# 表格每行的“数据要求/字符要求”会自动读成规则，这里只补充表格里没写的合理范围。
DEFAULT_RULES = [
    {"填报项目": "高度", "数据类型": "数值", "单位": "米", "取值范围": "100~9000"},
]

GENERAL_RULES = [
    "数据来源必须用书名号《》填写资料全称",
    "来源日期格式为 YYYY-MM-DD 或 YYYY-MM，须与资料上标注的发布日期一致",
    "每项数据必须针对“名称”所指的对象本身",
]


def rows_of(values):
    """把多个对象的数据按 ITEMS 顺序展开成表格行。"""
    rows = []
    for i, (content, source, pub) in enumerate(values):
        item, req, char_req = ITEMS[i % len(ITEMS)]
        rows.append([item, req, char_req, content, source, pub])
    return rows


def main():
    config.KB_DIR.mkdir(exist_ok=True)
    (config.KB_DIR / "青石县志（2018年版）.txt").write_text(COUNTY_ANNALS, encoding="utf-8")
    (config.KB_DIR / "青石县植被调查报告（2019年）.txt").write_text(VEGETATION_REPORT, encoding="utf-8")

    write_table(config.TEMPLATE_DIR / "山体信息表_空白模板.xlsx", "山体信息表",
                [[item, req, char_req] for item, req, char_req in ITEMS])
    write_table(config.SAMPLE_DIR / "样例_已正确填写.xlsx", "山体信息表", rows_of(QINGSHI_OK))
    write_table(config.INPUT_DIR / "甲单位填报.xlsx", "山体信息表（甲单位）", rows_of(UNIT_A))
    write_table(config.INPUT_DIR / "乙单位填报.xlsx", "山体信息表（乙单位）", rows_of(UNIT_B))
    save_rules(config.RULES_FILE, DEFAULT_COLUMNS, DEFAULT_RULES, GENERAL_RULES)
    print("模拟数据已生成：资料库 2 份资料、空白模板、正确样例、2 份待审核填报表、默认审核规则。")


if __name__ == "__main__":
    main()
