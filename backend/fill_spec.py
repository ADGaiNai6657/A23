import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation

CITY_TEMPLATE = "模板_2025中国城市经济百强.xlsx"
CITY_OUTPUT = "filled_2025中国城市经济百强.xlsx"
COVID_OUTPUT = "filled_covid19_byday.xlsx"
COVID_START = date(2020, 7, 1)
COVID_END = date(2020, 8, 31)


@dataclass(frozen=True)
class CityColumn:
    header: str
    key: str
    units: tuple[str, ...]


CITY_COLUMNS = (
    CityColumn("GDP总量（亿元）", "gdp", ("亿元",)),
    CityColumn("常住人口（万）", "population", ("万", "万人", "万人口")),
    CityColumn("人均GDP（元）", "gdp_per_capita", ("元", "元/人", "元／人")),
    CityColumn("一般公共预算收入（亿元）", "budget_revenue", ("亿元",)),
)
CITY_HEADERS = ("城市名", *(column.header for column in CITY_COLUMNS))

METRIC_ALIASES = {
    "gdp_per_capita": ("人均GDP", "人均地区生产总值", "人均生产总值"),
    "gdp": ("GDP总量", "地区生产总值", "生产总值", "GDP"),
    "population": ("常住人口", "人口基数", "人口严控至", "万人口", "人口"),
    "budget_revenue": ("一般公共预算收入", "公共预算收入"),
    "budget_expense": ("一般公共预算支出", "公共预算支出"),
    "sales": ("销售总额", "销售额"),
    "business_revenue": ("营业收入", "营收"),
    "retail_sales": ("社会消费品零售总额", "社零总额"),
    "new_cases": ("新增确诊病例", "新增病例", "new_cases"),
    "total_cases": ("累计确诊病例", "累计病例", "total_cases"),
    "new_deaths": ("新增死亡病例", "新增死亡", "new_deaths"),
    "total_deaths": ("累计死亡病例", "累计死亡", "total_deaths"),
}
COUNTRY_KEYS = ("location", "country", "countryregion", "countryname", "国家或地区", "国家地区", "国家", "国家名称", "国家或地区名称")
DATE_KEYS = ("date", "日期", "统计日期")


def normalize(value: str) -> str:
    return re.sub(r"[\s_（()）/／-]+", "", value).casefold()


ALIASES = {
    normalize(alias): key
    for key, aliases in METRIC_ALIASES.items()
    for alias in aliases
}
LONG_ALIASES = sorted(ALIASES, key=len, reverse=True)


def metric_key(metric: str) -> str:
    normalized = normalize(metric)
    if normalized in ALIASES:
        return ALIASES[normalized]
    for alias in LONG_ALIASES:
        if alias in normalized:
            return ALIASES[alias]
    return normalized


def number(value: object) -> Decimal | None:
    try:
        parsed = Decimal(re.sub(r"[\s,，]", "", str(value)))
    except InvalidOperation:
        return None
    return parsed if parsed.is_finite() else None


def choose_value(candidates: list[dict]) -> tuple[dict | None, str]:
    if not candidates:
        return None, "没有符合日期、指标和统计口径的来源"
    if len({candidate["value"] for candidate in candidates}) > 1:
        return None, "同口径来源存在不同值，缺少裁决依据"
    return candidates[0], "同口径候选值一致" if len(candidates) > 1 else "采用唯一符合口径的来源"
