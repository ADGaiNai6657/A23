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
CLAUSE_END = re.compile(r"[。！？；;、]|(?<!\d)[,，]|[,，](?!\d)")
PARTIAL_PERIOD = re.compile(
    r"第?[一二三四1234]季度|前[一二三123]季度|[上下]半年|前[一二三四五六七八九十0-9]+个?月"
    r"|(?:截至|截止到?)(?:(?:19|20)\d{2}年)?(?:[一二三四五六七八九十0-9]+月|第?[一二三四1234]季)"
    r"|(?<!\d)(?:1[0-2]|0?[1-9])月(?:份|末)?"
)
SCOPE_HEADING = re.compile(r"(?:(?:19|20)\d{2}年)?(?:" + PARTIAL_PERIOD.pattern + r")(?:数据|统计|情况)?(?:如下)?")


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
            if not NUMBER_TOKEN.search(gap):
                within_quote = (0 <= quoted_at <= quantity.start()
                                and quoted_at <= lo + label.start() and lo + label.end() <= quoted_at + len(quote))
                labels.append(((not within_quote, len(gap)), metric_key(label.group())))
        if not labels:
            matches.append((None, None))
            continue
        distance = min(gap for gap, _ in labels)
        local_keys = {key for gap, key in labels if gap == distance}
        key = next(iter(local_keys)) if len(local_keys) == 1 else None
        dated = PARTIAL_PERIOD.search(clause)
        if not dated and left:
            previous = text[left[-2].end() if len(left) > 1 else 0:left[-1].start()]
            previous = re.split(r"[：:]", previous.rstrip("：:"))[-1]
            dated = SCOPE_HEADING.fullmatch(previous)
        matches.append((key, dated.group() if dated else None))
    if not matches or any(key != metric_key(fact.metric) for key, _ in matches):
        return "数值对应的原文指标不匹配或无法确定", None
    return None, next((scope for _, scope in matches if scope), None)
