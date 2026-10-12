"""第 1 步：只给一个对象名称，为每个数据类型写出“数据要求”和“字符要求”，生成模板。

不用大模型时（默认），常见数据类型使用内置的默认要求，其余留空由你填写；
用大模型时（.env 里 USE_LLM=auto），由 AI 根据对象情况定制要求、消除歧义。

用法：
    python define.py 喜马拉雅山
    python define.py 喜马拉雅山 --items 名称,高度,地理坐标,地形特点,覆盖植被类型
    python define.py 喜马拉雅山 --llm      # 本次用大模型定制要求
生成 模板/<名称>_模板.xlsx。请打开检查、修改要求，确认后再运行 fill.py 填写。
"""
import argparse
import re

import config
from audit import llm_available
from kb import KnowledgeBase, search_all
from llm import ask_json
from rules import CHAR_TYPES
from table import DEFAULT_ITEMS, DEFAULT_REQUIREMENTS, write_table

SYSTEM_PROMPT = "你是数据填报规范的制定者，要写出明确、无歧义、可以逐条核查的填写要求。以 JSON 格式输出。"

USER_PROMPT = """要为对象【{name}】制作一张填报表，数据类型依次为：{items}。
请为每个数据类型写出“数据要求”和“字符要求”。

【资料库中与该对象相关的片段（可能为空，仅用于了解对象是什么）】
{context}

要求：
1. 先判断对象是什么（如单座山峰还是山脉、是否有多个同名对象），写在“对象说明”里。
2. 数据要求要消除歧义：说清楚指什么（例如山脉的“高度”是指最高峰海拔还是平均海拔、坐标是哪个点）、
   单位、精度、以哪类权威资料为准、同一数据有新旧版本时以哪个为准。
3. 字符要求只能用这些说法组合，便于程序自动检查：
   字符类型：{char_types}
   字符数：“6位”“2~10个字”“不超过100字”“至少2个字”
   描述性内容一般只写字数上限。
4. “名称”这一项的数据要求说明应填写的规范名称。

输出 JSON：
{{"对象说明": "一两句话",
  "项目": [{{"数据类型": "", "数据要求": "", "字符要求": ""}}]}}"""


def define_without_llm(name: str, items: list[str], has_material: bool):
    rows, unknown = [], []
    for item in items:
        req, char_req = DEFAULT_REQUIREMENTS.get(item, ("", ""))
        if not req:
            unknown.append(item)
        rows.append([item, req, char_req, name if item == "名称" else None])
    path = config.TEMPLATE_DIR / f"{name}_模板.xlsx"
    write_table(path, f"{name} 信息表", rows)
    print(f"已按内置默认要求为【{name}】生成模板（未使用大模型）：")
    for row in rows:
        print(f"  {row[0]}\n    数据要求：{row[1] or '（请填写）'}\n    字符要求：{row[2] or '（请填写）'}")
    if unknown:
        print(f"\n这些数据类型没有默认要求，请打开模板自己填写：{'、'.join(unknown)}")
    print("\n注意：默认要求是通用写法。例如对象是山脉时，请在“高度”里写明是最高峰海拔还是平均海拔。")
    if not has_material:
        print("提示：资料库里还没有该对象的资料，下一步填写时会全部留空。请先把资料放进 资料库/。")
    print(f"\n已生成 {path}\n请打开检查、修改要求，确认后运行：python fill.py \"{path}\"")


def main():
    parser = argparse.ArgumentParser(description="AI 定义数据要求和字符要求，生成填报模板")
    parser.add_argument("name", help="对象名称，例如 喜马拉雅山")
    parser.add_argument("--items", default=",".join(DEFAULT_ITEMS),
                        help="数据类型，用逗号分隔（默认：%(default)s）")
    parser.add_argument("--llm", action="store_true", help="本次用大模型定制要求（需要 API key）")
    args = parser.parse_args()
    items = [i for i in re.split(r"[,，、]", args.items) if i.strip()]
    if args.llm:
        config.USE_LLM = "auto"

    kb = KnowledgeBase()
    hits = search_all(kb, args.name, k=4)
    if not llm_available():
        define_without_llm(args.name, items, bool(hits))
        return
    context = "\n\n".join(f"（{doc.path.stem}）{chunk}" for doc, chunk in hits) or "（资料库中没有相关资料）"

    print(f"正在为【{args.name}】定义 {len(items)} 个数据类型的要求...")
    result = ask_json(SYSTEM_PROMPT, USER_PROMPT.format(
        name=args.name, items="、".join(items), context=context,
        char_types="、".join(CHAR_TYPES)))

    by_item = {r.get("数据类型"): r for r in result.get("项目", [])}
    rows = []
    for item in items:
        r = by_item.get(item, {})
        rows.append([item, r.get("数据要求", ""), r.get("字符要求", ""),
                     args.name if item == "名称" else None])

    path = config.TEMPLATE_DIR / f"{args.name}_模板.xlsx"
    write_table(path, f"{args.name} 信息表", rows)
    print(f"\n对象说明：{result.get('对象说明', '')}\n")
    for row in rows:
        print(f"  {row[0]}\n    数据要求：{row[1]}\n    字符要求：{row[2]}")
    if not hits:
        print("\n提示：资料库里还没有该对象的资料，下一步填写时会全部留空。请先把资料放进 资料库/。")
    print(f"\n已生成 {path}\n请打开检查、修改要求，确认后运行：python fill.py \"{path}\"")


if __name__ == "__main__":
    main()
