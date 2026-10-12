"""程序核查：不调用大模型，直接从资料原文里找数字、日期、坐标、名称并比较。

每个检查返回 Finding：
  支持 / 不支持 / 年份不符 —— 程序有把握的结论，附原文那一句作证据
  不确定                   —— 程序找不到或找到多个互相矛盾的候选，交给模型或人工
"""
import difflib
import re
from dataclasses import dataclass

from rules import SUBJECT_ITEMS, compare_coordinates, normalize_date, parse_coordinate, to_number


@dataclass
class Finding:
    结论: str               # 支持 / 不支持 / 年份不符 / 不确定
    说明: str = ""
    建议值: str | None = None
    证据: str = ""


# ---------- 公共工具 ----------

def sentences(text: str) -> list[str]:
    """按句号、分号、换行切句。"""
    return [s.strip() for s in re.split(r"(?<=[。！？；])|\n", text) if s and s.strip()]


NEW_OBJECT = re.compile(r"^[一-鿿]{2,8}(位于|地处|为|属|是)")


def subject_sentences(text: str, subject: str, follow: int = 2) -> list[str]:
    """与对象有关的句子：从提到对象的那句开始，到段尾，再加上后面最多 follow 段
    （资料常把一个对象的描述分成几段写，后面几段不再重复对象名）。
    遇到标题行、或以别的对象名开头的段落（如“白云岭位于……”）就停。没有对象名时返回全部句子。"""
    if not subject:
        return sentences(text)
    out, remaining = [], 0
    for para in (p.strip() for p in text.splitlines()):
        if not para:
            continue
        if subject in para:
            hit = False
            for s in sentences(para):
                hit = hit or subject in s
                if hit:
                    out.append(s)
            remaining = follow
        elif remaining:
            is_heading = len(para) < 20 and "。" not in para
            if is_heading or NEW_OBJECT.match(para):
                remaining = 0
                continue
            out += sentences(para)
            remaining -= 1
    return out


def _plain(s: str) -> str:
    return re.sub(r"[\s\W_]+", "", s)


# ---------- 数值 ----------

# 单位 -> (量纲, 换算到基本单位的倍数)
UNITS = {
    "米": ("长度", 1), "m": ("长度", 1), "千米": ("长度", 1000), "公里": ("长度", 1000), "km": ("长度", 1000),
    "平方米": ("面积", 1), "平方千米": ("面积", 1e6), "平方公里": ("面积", 1e6), "km²": ("面积", 1e6),
    "公顷": ("面积", 1e4), "亩": ("面积", 2000 / 3), "万亩": ("面积", 2e7 / 3),
    "人": ("人数", 1), "万人": ("人数", 1e4), "亿人": ("人数", 1e8),
    "元": ("金额", 1), "万元": ("金额", 1e4), "亿元": ("金额", 1e8),
    "千克": ("质量", 1), "公斤": ("质量", 1), "吨": ("质量", 1e3), "万吨": ("质量", 1e7),
}
UNIT_RE = "|".join(sorted(map(re.escape, UNITS), key=len, reverse=True))
RANGE_WORDS = re.compile(r"\s*(以下|以上|以内|之间|左右|至|到|~|～|—|-)")

# 数据类型里出现左边的词时，资料里也认右边这些说法
SYNONYMS = {
    "高度": ["海拔", "高程", "高度"], "海拔": ["海拔", "高程"], "高程": ["高程", "海拔"],
    "长度": ["全长", "长度"], "全长": ["全长"], "面积": ["面积"], "人口": ["人口"],
    "水深": ["水深"], "深度": ["水深", "深度"], "产量": ["产量"], "生产总值": ["生产总值", "GDP"],
}


def target_unit(row: dict, rule: dict) -> str:
    unit = rule.get("单位", "")
    if not unit and (m := re.search(r"单位(?:为|是|[：:])\s*([^\s，,；;。（(）)]+)", row.get("要求", ""))):
        unit = m[1]
    return unit if unit in UNITS else ""


def keyword_tiers(row: dict) -> list[list[str]]:
    """先用最具体的说法（数据类型去掉年份），找不到再用同义词。"""
    item = re.sub(r"^\d{4}年(末|底|初)?", "", row["项目"]).strip()
    specific = [item] if len(item) >= 2 else []
    general = []
    text = row["项目"] + " " + row.get("要求", "").split("，")[0]
    for key, words in SYNONYMS.items():
        if key in text:
            general += [w for w in words if w not in general and w not in specific]
    return [t for t in (specific, sorted(general, key=len, reverse=True)) if t]


def find_numbers(row: dict, rule: dict, text: str) -> tuple[list, list]:
    """返回 (年份对得上的候选, 年份对不上的候选)，候选为 (换算到目标单位的数值, 原文数字串, 句子)。"""
    unit = target_unit(row, rule)
    item_years = set(re.findall(r"(\d{4})年", row["项目"]))
    sents = subject_sentences(text, row.get("主体", ""))
    for tier in keyword_tiers(row):
        good, wrong_year = [], []
        for s in sents:
            for kw in tier:
                pattern = re.escape(kw) + r"[^。；，,\d]{0,8}?(\d[\d,]*(?:\.\d+)?)\s*(" + UNIT_RE + r")?"
                for m in re.finditer(pattern, s):
                    if RANGE_WORDS.match(s, m.end()):
                        continue                               # “海拔900米以下”是范围界线，不是数值本身
                    value = to_number(m[1])
                    src_unit = m[2] or ""
                    if value is None:
                        continue
                    if unit and src_unit:
                        if UNITS[src_unit][0] != UNITS[unit][0]:
                            continue                           # 量纲不同，不是同一个量
                        value = value * UNITS[src_unit][1] / UNITS[unit][1]
                    sent_years = set(re.findall(r"(\d{4})年", s))
                    bucket = wrong_year if item_years and sent_years and not item_years & sent_years else good
                    bucket.append((value, m[1], s))
        if good or wrong_year:
            return good, wrong_year
    return [], []


def _decimals(s: str) -> int:
    return len(s.split(".")[1]) if "." in s else 0


def format_number(value: float, decimals: int) -> str:
    return f"{value:.{decimals}f}" if decimals > 0 else str(int(round(value)))


def check_number(row: dict, rule: dict, text: str) -> Finding:
    good, wrong_year = find_numbers(row, rule, text)
    filled = to_number(row["内容"].replace("约", ""))
    if not good:
        if wrong_year:
            s = wrong_year[0][2]
            return Finding("年份不符", f"资料里找到的数据年份与“{row['项目']}”要求的年份不一致", 证据=s)
        return Finding("不确定", "程序没在资料里找到这个数据的明确写法")
    distinct = {round(v, 6) for v, _, _ in good}
    if len(distinct) > 1:
        values = "、".join(sorted({raw for _, raw, _ in good}))
        return Finding("不确定", f"资料里有多个候选数值（{values}），程序无法确定是哪个", 证据=good[0][2])
    value, raw, sent = good[0]
    rule_digits = rule.get("小数位", "")
    digits = int(rule_digits) if rule_digits.isdigit() else _decimals(raw)
    suggestion = format_number(value, digits)
    if filled is None:
        return Finding("不支持", f"填写的不是数字，资料中为 {raw}", 建议值=suggestion, 证据=sent)
    tolerance = 0.5 * 10 ** -max(_decimals(row["内容"]), 0) + 1e-9
    if abs(filled - value) <= tolerance:
        return Finding("支持", "与资料数值一致", 证据=sent)
    return Finding("不支持", f"资料中为 {suggestion}，填写为 {row['内容']}", 建议值=suggestion, 证据=sent)


# ---------- 坐标 ----------

COORD_TEXT = re.compile(r"(?:东经|西经|[EW])\s*[\d.°′″'\"\s]+[、，,\s]*(?:北纬|南纬|[NS])\s*[\d.°′″'\"]+")


def find_coordinates(row: dict, text: str) -> list[tuple]:
    """返回 [(十进制坐标, 原文坐标串, 句子)]。"""
    found = []
    for s in subject_sentences(text, row.get("主体", "")):
        if (m := COORD_TEXT.search(s)) and (p := parse_coordinate(m[0])):
            found.append((p, m[0], s))
    return found


def check_coordinate(row: dict, text: str) -> Finding:
    found = find_coordinates(row, text)
    distinct = {(round(p[0], 4), round(p[1], 4)) for p, _, _ in found}
    if not found:
        return Finding("不确定", "程序没在资料里找到该对象的坐标")
    if len(distinct) > 1:
        return Finding("不确定", "资料里有多个坐标，程序无法确定是哪个", 证据=found[0][2])
    _, coord_text, sent = found[0]
    verdict = compare_coordinates(row["内容"], coord_text)
    suggestion = coord_text.replace("、", "，")
    if verdict == "一致":
        return Finding("支持", "换算后与资料坐标一致（误差不超过 1′）", 证据=sent)
    if verdict == "经纬度写反":
        return Finding("不支持", "经度和纬度写反了", 建议值=suggestion, 证据=sent)
    if verdict == "不一致":
        return Finding("不支持", "换算后与资料坐标相差超过 1′", 建议值=suggestion, 证据=sent)
    return Finding("不确定", "填写的坐标格式程序认不出")


# ---------- 名称、文字 ----------

def phrase_found(phrase: str, sents: list[str], ratio: float = 0.8) -> str:
    """短语是否在某一句里出现：要求连续相同的部分至少占短语的 80%，避免把“东坡陡峭”当成“西坡陡峭”。
    找到时返回那一句，否则返回空字符串。"""
    p = _plain(phrase)
    for s in sents:
        q = _plain(s)
        if p in q:
            return s
        m = difflib.SequenceMatcher(None, p, q, autojunk=False).find_longest_match(0, len(p), 0, len(q))
        if m.size >= max(2, ratio * len(p)):
            return s
    return ""


# 资料在“引用别人的说法”或“辟谣”时常见的字眼：原文出现了相同的话，也不代表资料认可它
QUOTE_OR_REFUTE = re.compile(
    "网传|網傳|传言|傳言|流传|流傳|谣言|謠言|辟谣|闢謠|不实|不實|假消息|误传|誤傳|误解|誤解|"
    "事实查核|事實查核|查核|核实|核實|并非|並非|并不|並不|不是真的|是真的吗|是真的嗎|没有科学依据|沒有科學依據|"
    "所谓|所謂|据说|據說|有人说|有人說|有人认为|有人認為|转发|轉發|转贴|轉貼|朋友圈|"
    "实际上|實際上|事实上|事實上|其实|其實|真相|误区|誤區|错误|錯誤|迷思|谣传|謠傳|传闻|傳聞|"
    "[？?]|[「」“”『』]")


def quoted_or_refuted(sents: list[str], index: int) -> str:
    """看这句和前一句、后两句里有没有“引用 / 辟谣”的字眼，返回命中的字眼（没有则返回空字符串）。"""
    window = "".join(sents[max(0, index - 1): index + 3])
    m = QUOTE_OR_REFUTE.search(window)
    return m[0] if m else ""


def check_text(row: dict, text: str) -> Finding:
    content = row["内容"]
    all_sents = sentences(text)
    if content and _plain(content) in _plain(text):
        index = next((i for i, s in enumerate(all_sents) if _plain(content) in _plain(s)), None)
        if index is None:     # 跨句出现：用开头几个字定位到起始句
            head = _plain(content)[:8]
            index = next((i for i, s in enumerate(all_sents) if head in _plain(s)), None)
            if index is None:
                return Finding("不确定", "资料中有相同的文字，但跨越多句，程序无法确认上下文，请人工确认")
        sent = all_sents[index]
        if cue := quoted_or_refuted(all_sents, index):
            return Finding("不确定", f"资料中出现了相同的说法，但上下文像是在引用或辟谣（出现“{cue}”），请人工确认",
                           证据="".join(all_sents[max(0, index - 1): index + 3])[:200])
        return Finding("支持", "资料中有相同的写法", 证据=sent)
    if row["项目"] in SUBJECT_ITEMS:
        return Finding("不确定", f"资料中没有出现“{content}”这个名称（可能是别名或写法不同）")
    # 描述性内容：拆成短语逐个到原文里找
    sents = subject_sentences(text, row.get("主体", "")) or sentences(text)
    phrases = [p for p in re.split(r"[，,、；;。\s和及与]+", content) if len(_plain(p)) >= 2]
    if not phrases:
        return Finding("不确定", "描述性内容，程序无法判断意思是否一致")
    hits = {p: phrase_found(p, sents) for p in phrases}
    missing = [p for p, s in hits.items() if not s]
    evidence = "".join(dict.fromkeys(s for s in hits.values() if s))[:200]
    if not missing:
        for s in dict.fromkeys(hits.values()):
            index = sents.index(s)
            if cue := quoted_or_refuted(sents, index):
                return Finding("不确定", f"各项说法在资料中都能找到，但上下文像是在引用或辟谣（出现“{cue}”），"
                                         "请人工确认", 证据=evidence)
        return Finding("支持", "各项说法在资料中都能找到", 证据=evidence)
    if len(missing) == len(phrases):
        return Finding("不确定", "资料中没有找到这些说法：" + "、".join(missing), 证据=evidence)
    return Finding("不确定", "资料中找到了部分说法，没找到：" + "、".join(missing), 证据=evidence)


def check_content(row: dict, rule: dict, text: str) -> Finding:
    if rule.get("字符类型") == "经纬度":
        return check_coordinate(row, text)
    if rule.get("数据类型") == "数值":
        return check_number(row, rule, text)
    return check_text(row, text)


def related_text(row: dict, text: str, limit: int = 200) -> str:
    """给人工复核看的相关原文：优先取提到对象且含数据类型关键词的句子。"""
    sents = subject_sentences(text, row.get("主体", "")) or sentences(text)
    item = row["项目"]
    words = [w for tier in keyword_tiers(row) for w in tier] + [item[i:i + 2] for i in range(len(item) - 1)]
    hits = [s for s in sents if any(w in s for w in words)] or sents[:2]
    return "".join(hits)[:limit]


# ---------- 来源日期 ----------

DATE_IN_TEXT = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?|(\d{4})[-./](\d{1,2})(?:[-./](\d{1,2}))?")
PUBLISH_WORDS = re.compile(r"出版|发布|印发|公布|发行|编|成文")


def find_publish_date(text: str, head_chars: int = 300) -> tuple[tuple | None, str]:
    """在资料开头找发布日期，返回 ((年, 月, 日或None), 所在那一行)；找不到或有多个时返回 (None, 说明)。"""
    head = text[:head_chars]
    found = []
    for m in DATE_IN_TEXT.finditer(head):
        y, mo, d = (m[1], m[2], m[3]) if m[1] else (m[4], m[5], m[6])
        end = head.find("\n", m.end())
        line = head[head.rfind("\n", 0, m.start()) + 1: end if end != -1 else len(head)]
        near = head[max(0, m.start() - 8): m.end() + 8]
        found.append(((int(y), int(mo), int(d) if d else None), line.strip(), bool(PUBLISH_WORDS.search(near))))
    if not found:
        return None, "资料开头没有找到日期"
    preferred = [f for f in found if f[2]] or found
    distinct = {f[0] for f in preferred}
    if len(distinct) > 1:
        return None, "资料开头有多个日期，程序无法确定哪个是发布日期"
    return preferred[0][0], preferred[0][1]


def format_date(date: tuple) -> str:
    y, mo, d = date
    return f"{y}-{mo:02d}" + (f"-{d:02d}" if d else "")


def extract_value(row: dict, rule: dict, text: str) -> tuple[str, str] | None:
    """填表用：从一份资料里提取这一项的值，返回 (值, 原文句子)；提取不到或有歧义时返回 None。"""
    if rule.get("字符类型") == "经纬度":
        found = find_coordinates(row, text)
        if len({(round(p[0], 4), round(p[1], 4)) for p, _, _ in found}) == 1:
            return found[0][1].replace("、", "，"), found[0][2]
        return None
    if rule.get("数据类型") == "数值":
        good, _ = find_numbers(row, rule, text)
        if len({round(v, 6) for v, _, _ in good}) == 1:
            value, raw, sent = good[0]
            digits = rule.get("小数位", "")
            return format_number(value, int(digits) if digits.isdigit() else _decimals(raw)), sent
        return None
    if row["项目"] in SUBJECT_ITEMS and row["内容"] and row["内容"] in text:
        return row["内容"], next((s for s in sentences(text) if row["内容"] in s), "")
    return None


def check_date(filled: str, text: str) -> Finding:
    filled = normalize_date(filled) or filled     # “2018-6”等写法先统一，格式问题由格式检查单独报
    m = re.fullmatch(r"(\d{4})-(\d{2})(?:-(\d{2}))?", filled or "")
    if not m:
        return Finding("不确定", "填写的日期格式程序认不出")
    date, line = find_publish_date(text)
    if date is None:
        return Finding("不确定", line)
    y, mo, d = date
    fy, fm, fd = int(m[1]), int(m[2]), int(m[3]) if m[3] else None
    same = (fy, fm) == (y, mo) and (fd is None or d is None or fd == d)
    if same:
        return Finding("支持", "与资料标注的日期一致", 证据=line)
    suggestion = f"{y}-{mo:02d}" + (f"-{d:02d}" if fd is not None and d is not None else "")
    return Finding("不支持", f"资料标注的日期为 {suggestion}，填写为 {filled}", 建议值=suggestion, 证据=line)
