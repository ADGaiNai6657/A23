import re

from extractor import NUMBER_TOKEN, compact, quoted_quantities
from fill_spec import METRIC_ALIASES, metric_key
from schema import RawFact

METRIC_LABEL = re.compile(
    "|".join(re.escape(label) for label in sorted(
        {compact(alias) for key in ("gdp", "population", "gdp_per_capita", "budget_revenue", "budget_expense")
         for alias in METRIC_ALIASES[key]}, key=lambda label: (-len(label), label),
    )),
    re.IGNORECASE,
)
# A comma inside a thousands group is part of the number.
CLAUSE_END = re.compile(r"[。！？；;、]|(?<!\d)[,，]|[,，](?!\d)|其中")
RATE_OR_SHARE = re.compile(r"增速|增长率|变动率|涨跌幅|增幅|降幅|占比|比重|百分点|增量|减量|差额")
AMOUNT_CHANGE = re.compile(r"(?:增加|减少|增长|下降|上升|下滑|下跌|提高|提升|降低|回落)(?!到|至|为)")
RELATIVE = re.compile(r"同比|环比|(?:较|比)(?:上年|上月|上季|同期|上期|去年)")
LEVEL_TARGET = re.compile(r"(?:增加|减少|增长|下降|上升|下滑|提高|提升|降低|回落)(?:到|至|为)")
MONTH = r"(?:1[0-2]|0?[1-9]|十[一二]?|[一二三四五六七八九])"
MONTH_RANGE = rf"{MONTH}月?(?:至|到|[-—–~～]){MONTH}月"
QUARTER = r"(?:第?[一二三四1234]季(?:度)?|前[一二三123]季(?:度)?)"
PARTIAL_PERIOD = re.compile(
    rf"(?:截至|截止到?)(?:(?:19|20)\d{{2}}年)?(?:{MONTH}月|{QUARTER})"
    rf"|{QUARTER}|[上下]半年|前{MONTH}个?月"
    rf"|(?<!\d)(?:{MONTH_RANGE}|{MONTH}月)(?:份|末)?"
)
FULL_YEAR = re.compile(r"全年(?:度)?")
YEAR_LABEL = re.compile(r"(?:19|20)\d{2}年")
SCOPE_HEADING = re.compile(r"(?:(?:19|20)\d{2}年)?(?:" + PARTIAL_PERIOD.pattern + r")(?:数据|统计|情况)?(?:如下)?")
PARENTHESIS = re.compile(r"[（()）]")


def has_quantity(text: str) -> bool:
    return NUMBER_TOKEN.search(YEAR_LABEL.sub("", PARTIAL_PERIOD.sub("", text))) is not None


def quantity_scope(clause: str, start: int, end: int) -> str | None:
    """Keep a parenthetical subquantity's period with that subquantity."""
    openings, annotations = [], []
    for mark in PARENTHESIS.finditer(clause):
        if mark.group() in "（(":
            openings.append(mark.start())
        elif openings:
            annotations.append((openings.pop(), mark.end()))

    covering = next(((left, right) for left, right in annotations if left <= start and end <= right), None)
    if covering:
        left, right = covering
        nested = quantity_scope(clause[left + 1:right - 1], start - left - 1, end - left - 1)
        if nested or "全年" in clause[left + 1:right - 1]:
            return nested

    local = list(clause)
    for left, right in annotations:
        if left < end and start < right:
            continue
        if has_quantity(clause[left:right]):
            local[left:right] = " " * (right - left)

    context = "".join(local)
    trailing = context[end:]
    if not has_quantity(trailing) and (postfix := PARTIAL_PERIOD.search(trailing)):
        return postfix.group()
    prefix = context[:start]
    partial = list(PARTIAL_PERIOD.finditer(prefix))
    annual = list(FULL_YEAR.finditer(prefix))
    if partial and (not annual or partial[-1].start() > annual[-1].start()):
        return partial[-1].group()
    return None


def is_change(gap: str) -> bool:
    return bool(RATE_OR_SHARE.search(gap) or AMOUNT_CHANGE.search(gap)
                or (RELATIVE.search(gap) and not LEVEL_TARGET.search(gap)))


def city_claim(fact: RawFact, source: str) -> tuple[str | None, str | None]:
    """Resolve the number's local indicator and any explicitly partial reporting period."""
    text = compact(source)
    quote = compact(fact.snippet)
    separators = list(CLAUSE_END.finditer(text))
    matches = []
    for quantity in quoted_quantities(fact, text):
        left = [mark for mark in separators if mark.end() <= quantity.start()]
        lo = left[-1].end() if left else 0
        hi = next((mark.start() for mark in separators if mark.start() >= quantity.end()), len(text))
        clause = text[lo:hi]
        start, end = quantity.start() - lo, quantity.end() - lo
        labels = []
        quoted_at = text.find(quote, max(0, quantity.end() - len(quote)))
        for label in METRIC_LABEL.finditer(clause):
            if label.end() <= start:
                gap = clause[label.end():start]
            elif label.start() >= end:
                gap = clause[end:label.start()]
            else:
                continue
            if not has_quantity(gap):
                within_quote = (0 <= quoted_at <= quantity.start()
                                and quoted_at <= lo + label.start() and lo + label.end() <= quoted_at + len(quote))
                labels.append(((not within_quote, len(gap)), metric_key(label.group()), is_change(gap)))
        distance = min((gap for gap, _, _ in labels), default=None)
        local_claims = {(key, change) for gap, key, change in labels if gap == distance}
        key, change = next(iter(local_claims)) if len(local_claims) == 1 else (None, False)
        scope = quantity_scope(clause, start, end)
        if not scope and left:
            previous = text[left[-2].end() if len(left) > 1 else 0:left[-1].start()]
            previous = re.split(r"[：:]", previous.rstrip("：:"))[-1]
            dated = SCOPE_HEADING.fullmatch(previous)
            scope = dated.group() if dated else None
        matches.append((key, change, scope))
    scope = next((scope for _, _, scope in matches if scope), None)
    if not matches or any(key != metric_key(fact.metric) for key, _, _ in matches):
        return "数值对应的原文指标不匹配或无法确定", scope
    if any(change for _, change, _ in matches):
        return "原文数值为增减量、增速或比率，不能用于总量或水平字段", scope
    return None, scope
