import re

from extractor import NUMBER_TOKEN, compact, quoted_quantities
from fill_spec import CITY_COLUMNS, METRIC_ALIASES, metric_key
from schema import RawFact

METRIC_LABEL = re.compile(
    "|".join(re.escape(label) for label in sorted(
        {compact(alias) for key in ("gdp", "population", "gdp_per_capita", "budget_revenue", "budget_expense")
         for alias in METRIC_ALIASES[key]}, key=lambda label: (-len(label), label),
    )),
    re.IGNORECASE,
)
SENTENCE_END = "。！？；;!?"
# A comma inside a thousands group is part of the number.
CLAUSE_END = re.compile(rf"[{SENTENCE_END}、]|(?<!\d)[,，]|[,，](?!\d)|其中")
DERIVED_MEASURE = re.compile(r"净增|净减|新增|增速|增长率|变动率|涨跌幅|增幅|降幅|占比|比重|百分点|增量|减量|差额")
DERIVED_PREFIX = re.compile(r"(?:净增(?:加)?|净减(?:少)?|新增|增量|减量|增速)$")
CHANGE_VERB = r"增加|减少|增长|下降|上升|下滑|下跌|提高|提升|降低|回落"
AMOUNT_CHANGE = re.compile(rf"(?:{CHANGE_VERB})(?!到|至|为)")
RELATIVE = re.compile(r"同比|环比|(?:较|比)(?:上年|上月|上季|同期|上期|去年|(?:19|20)\d{2}年)")
LEVEL_TARGET = rf"(?:{CHANGE_VERB}|调整|变动|转|变|增|升|降)(?:到|至|为)"
AMOUNT_UNIT = "|".join(re.escape(unit) for unit in sorted(
    {unit for column in CITY_COLUMNS for unit in column.units} | {"万亿元", "万元"},
    key=lambda unit: (-len(unit), unit),
))
TRANSITION = re.compile(
    rf"(?:从|由)(?:约)?(?P<origin>{NUMBER_TOKEN.pattern})(?:{AMOUNT_UNIT})?"
    rf"(?:同比|环比)?{LEVEL_TARGET}(?:约)?(?P<destination>{NUMBER_TOKEN.pattern})"
)
AMOUNT_RANGE = re.compile(
    rf"(?P<origin>{NUMBER_TOKEN.pattern})(?:{AMOUNT_UNIT})?(?:至|到|[-—–~～])"
    rf"(?P<destination>{NUMBER_TOKEN.pattern})"
)
MONTH = r"(?:1[0-2]|0?[1-9]|十[一二]?|[一二三四五六七八九])"
DAY = r"(?:3[01]|[12]\d|0?[1-9]|[一二三]?十[一二三四五六七八九]?|[一二三四五六七八九])"
MONTH_DAY = rf"{MONTH}月(?:(?:{DAY})[日号]|[底末])?"
MONTH_RANGE = rf"{MONTH}月?(?:至|到|[-—–~～]){MONTH}月"
QUARTER = r"(?:第?[一二三四1234]季(?:度)?|前[一二三123]个?季(?:度)?|首季(?:度)?)"
YEAR_LABEL = re.compile(r"(?:19|20)\d{2}年")
PARTIAL_PERIOD = re.compile(
    rf"(?:截至|截止(?:到|至)?)(?:{YEAR_LABEL.pattern})?(?:{MONTH_DAY}|{QUARTER})"
    rf"|{QUARTER}|[上下]半年|前{MONTH}个?月"
    rf"|(?<!\d)(?:{MONTH_RANGE}|{MONTH_DAY})(?:份|末)?"
)
FULL_YEAR = re.compile(r"全年(?:度)?")
REPORTING_PERIOD = re.compile(rf"{PARTIAL_PERIOD.pattern}|{FULL_YEAR.pattern}")
SCOPE_HEADING = re.compile(
    rf"(?:{YEAR_LABEL.pattern})?(?:按(?:照)?|(?:统计)?口径(?:为|是)?)?"
    rf"(?P<period>{REPORTING_PERIOD.pattern})(?:累计|数据|统计(?:数据)?|情况|口径)?(?:如下)?"
)
REPORTING_ASIDE = re.compile(
    r"(?:(?:经|据|根据|按(?:照)?)?(?:初步|最终)?(?:核算|统计|调查|测算|汇总)(?:结果|数据|显示)?"
    r"|经初步核算口径调整后|分别|主要指标如下)"
)
PARENTHESIS = re.compile(r"[（()）]")
NON_ACTUAL = re.compile(r"(?P<goal>目标(?!完成率|实现率)|力争|计划(?=为|达|实现|完成|全年))"
                        r"|(?P<forecast>预计|(?<!低于)(?<!超出)(?<!高于)预期|预测|有望)")
ACTUAL_MARKER = re.compile(r"实际|实绩|已实现|经(?:初步)?核算")
POSTFIX_ROLE = re.compile(
    r"(?:万亿元|亿元|万元|万人|元|%|％)?"
    r"(?:[（(](?P<bracket>目标值|预测值)[）)]|(?P<bare>目标值|预测值)(?!仅作比较))"
)
CLAUSE_NUMBER = re.compile(rf"{REPORTING_PERIOD.pattern}|{YEAR_LABEL.pattern}|(?P<amount>{NUMBER_TOKEN.pattern})")
CLAIM_ISSUES = {
    "goal": ("goal", "原文数值为目标，不能用于实绩字段"),
    "forecast": ("forecast", "原文数值为预测，不能用于实绩字段"),
    "origin": ("metric_mismatch", "原文数值为起点，不能用于总量或水平字段"),
    "range": ("metric_mismatch", "原文数值为区间端点，不能用于总量或水平字段"),
    "change": ("metric_mismatch", "原文数值为增减量、增速或比率，不能用于总量或水平字段"),
}


def has_quantity(text: str) -> bool:
    return NUMBER_TOKEN.search(YEAR_LABEL.sub("", PARTIAL_PERIOD.sub("", text))) is not None


def parentheses(clause: str) -> list[tuple[int, int]]:
    openings, annotations = [], []
    for mark in PARENTHESIS.finditer(clause):
        if mark.group() in "（(":
            openings.append(mark.start())
        elif openings:
            annotations.append((openings.pop(), mark.end()))
    return annotations


def has_confirmed_actual(text: str) -> bool:
    context = list(text)
    for opening, closing in parentheses(text):
        if "待核" in text[opening:closing]:
            context[opening:closing] = " " * (closing - opening)
    return ACTUAL_MARKER.search("".join(context)) is not None


def quantity_scope(clause: str, start: int, end: int) -> str | None:
    """Keep subordinate notes and scope qualifiers with the quantities they describe."""
    annotations = parentheses(clause)
    covering = next(((left, right) for left, right in annotations if left <= start and end <= right), None)
    if covering:
        left, right = covering
        nested = quantity_scope(clause[left + 1:right - 1], start - left - 1, end - left - 1)
        if nested:
            return nested

    local = list(clause)
    for left, right in annotations:
        if left < end and start < right:
            continue
        note = clause[left + 1:right - 1]
        if has_quantity(note) or (REPORTING_PERIOD.search(note) and not SCOPE_HEADING.fullmatch(note)):
            local[left:right] = " " * (right - left)

    context = "".join(local)
    trailing = context[end:]
    if not has_quantity(trailing) and (postfix := REPORTING_PERIOD.search(trailing)):
        return postfix.group()
    prefix = list(REPORTING_PERIOD.finditer(context, 0, start))
    return prefix[-1].group() if prefix else None


def prefix_non_actual_role(clause: str, start: int, end: int) -> str | None:
    # Earlier quantities can belong to a different claim in the same clause.
    left = max((item.end() for item in CLAUSE_NUMBER.finditer(clause, 0, start)
                if item.lastgroup == "amount"), default=0)
    prefix = list(clause[left:end])
    for opening, closing in parentheses(clause):
        if opening < end and closing <= end:
            prefix[max(0, opening - left):closing - left] = " " * (closing - max(left, opening))
    candidates = list(NON_ACTUAL.finditer("".join(prefix)))
    actual = list(ACTUAL_MARKER.finditer("".join(prefix)))
    if candidates and (not actual or candidates[-1].start() > actual[-1].start()):
        return candidates[-1].lastgroup
    return None


def non_actual_role(clause: str, start: int, end: int) -> str | None:
    role = prefix_non_actual_role(clause, start, end)
    if role:
        return role
    after = POSTFIX_ROLE.match(clause[end:])
    if after:
        return "goal" if (after["bracket"] or after["bare"]) == "目标值" else "forecast"
    return None


def reporting_year(clause: str) -> str | None:
    context = list(clause)
    for opening, closing in parentheses(clause):
        note = clause[opening:closing]
        if "比较" in note or re.search(r"(?:与|较|比)(?:19|20)\d{2}年.*(?:增加|减少|增长|下降|增减|高于|低于)", note):
            context[opening:closing] = " " * (closing - opening)
    years = list(YEAR_LABEL.finditer("".join(context)))
    return years[-1].group() if years else None


def shared_non_actual_role(text: str, left: list[re.Match[str]], clause: str) -> str | None:
    if has_confirmed_actual(clause):
        return None
    year = reporting_year(clause)
    for index in range(len(left) - 1, -1, -1):
        mark = left[index]
        if mark.group() in SENTENCE_END:
            break
        start = left[index - 1].end() if index else 0
        previous = text[start:mark.start()]
        previous_year = reporting_year(previous)
        if year and previous_year and year != previous_year:
            break
        year = year or previous_year
        if has_confirmed_actual(previous):
            break
        amounts = [item for item in CLAUSE_NUMBER.finditer(previous) if item.lastgroup == "amount"]
        if amounts:
            amount = amounts[-1]
            # A leading role belongs to the coordinated list; a postfixed label
            # belongs only to this amount, even when both appear together.
            role = prefix_non_actual_role(previous, amount.start(), amount.end())
            if role:
                return role
    return None


def inherited_scope(text: str, left: list[re.Match[str]], clause: str) -> str | None:
    """Walk back through a reporting sentence's indicators and neutral asides."""
    year = reporting_year(clause)
    for index in range(len(left) - 1, -1, -1):
        mark = left[index]
        if mark.group() in SENTENCE_END:
            break
        start = left[index - 1].end() if index else 0
        previous = re.split(r"[：:]", text[start:mark.start()].rstrip("：:"))[-1]
        # Notes within one indicator do not set the period of the list.
        context = list(previous)
        for opening, closing in parentheses(previous):
            context[opening:closing] = " " * (closing - opening)
        previous = "".join(context).strip()
        previous_year = reporting_year(previous)
        if year and previous_year and year != previous_year:
            break
        year = year or previous_year
        if heading := SCOPE_HEADING.fullmatch(previous):
            return heading["period"]
        if METRIC_LABEL.search(previous):
            amounts = [item for item in CLAUSE_NUMBER.finditer(previous) if item.lastgroup == "amount"]
            if amounts:
                amount = amounts[-1]
                if not non_actual_role(previous, amount.start(), amount.start()):
                    scope = quantity_scope(previous, *amount.span())
                    if scope:
                        return scope
                continue
        if not (REPORTING_ASIDE.fullmatch(previous) or YEAR_LABEL.fullmatch(previous)):
            break
    return None


def is_change(gap: str) -> bool:
    return bool(DERIVED_MEASURE.search(gap) or AMOUNT_CHANGE.search(gap)
                or (RELATIVE.search(gap) and not re.search(LEVEL_TARGET, gap)))


def quantity_claim_span(clause: str, start: int, end: int) -> tuple[int, int, str | None]:
    """Link a compound quantity to its indicator and mark an unusable origin or range bound."""
    for expression in (TRANSITION, AMOUNT_RANGE):
        for match in expression.finditer(clause):
            if (start, end) not in (match.span("origin"), match.span("destination")):
                continue
            role = "range" if expression is AMOUNT_RANGE else None
            if expression is TRANSITION and (start, end) == match.span("origin"):
                role = "origin"
            return match.start(), match.end(), role
    return start, end, None


def city_claim(fact: RawFact, source: str) -> tuple[tuple[str, str] | None, str | None]:
    """Resolve the original indicator, quantity eligibility and reporting period."""
    text = compact(source)
    quote = compact(fact.snippet)
    separators = list(CLAUSE_END.finditer(text))
    annotations = parentheses(text)
    matches = []
    for quantity in quoted_quantities(fact, text):
        boundaries = [mark for mark in separators if not any(
            annotation_start < mark.start() < annotation_end
            and not annotation_start <= quantity.start() < annotation_end
            for annotation_start, annotation_end in annotations
        )]
        left = [mark for mark in boundaries if mark.end() <= quantity.start()]
        lo = left[-1].end() if left else 0
        hi = next((mark.start() for mark in boundaries if mark.start() >= quantity.end()), len(text))
        clause = text[lo:hi]
        start, end = quantity.start() - lo, quantity.end() - lo
        claim_start, claim_end, compound_role = quantity_claim_span(clause, start, end)
        labels = []
        quoted_at = text.find(quote, max(0, quantity.end() - len(quote)))
        for label in METRIC_LABEL.finditer(clause):
            if label.end() <= claim_start:
                gap = clause[label.end():claim_start]
            elif label.start() >= claim_end:
                gap = clause[claim_end:label.start()]
            else:
                continue
            if not has_quantity(gap):
                within_quote = (0 <= quoted_at <= quantity.start()
                                and quoted_at <= lo + label.start() and lo + label.end() <= quoted_at + len(quote))
                role = (non_actual_role(clause, claim_start, end)
                        or shared_non_actual_role(text, left, clause) or compound_role)
                if role is None and (is_change(gap) or DERIVED_PREFIX.search(clause[:label.start()])
                                     or clause.startswith(("%", "％", "个百分点"), end)):
                    role = "change"
                labels.append(((not within_quote, len(gap)), metric_key(label.group()), role))
        distance = min((gap for gap, _, _ in labels), default=None)
        local_claims = {(key, role) for gap, key, role in labels if gap == distance}
        key, role = next(iter(local_claims)) if len(local_claims) == 1 else (None, None)
        scope = quantity_scope(clause, start, end) or inherited_scope(text, left, clause)
        matches.append((key, role, scope))
    scopes = [scope for _, _, scope in matches if scope]
    scope = next((scope for scope in scopes if not FULL_YEAR.fullmatch(scope)), scopes[0] if scopes else None)
    if not matches or any(key != metric_key(fact.metric) for key, _, _ in matches):
        return ("metric_mismatch", "数值对应的原文指标不匹配或无法确定"), scope
    roles = {role for _, role, _ in matches}
    return next((issue for role, issue in CLAIM_ISSUES.items() if role in roles), None), scope
