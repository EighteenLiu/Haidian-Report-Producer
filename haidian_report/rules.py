from __future__ import annotations

from collections import Counter
from datetime import datetime
from typing import Any


def fmt_int(value: Any) -> str:
    try:
        iv = int(round(float(value)))
        return str(iv)
    except Exception:
        return str(value if value is not None else "")


def fmt_float(value: Any, digits: int = 2) -> str:
    try:
        return f"{float(value):.{digits}f}"
    except Exception:
        return str(value if value is not None else "")


def fmt_pct(value: Any, digits: int = 2) -> str:
    try:
        return f"{float(value) * 100:.{digits}f}%"
    except Exception:
        text = str(value if value is not None else "")
        return text if text.endswith("%") else text


def fmt_datetime(value: Any) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value if value is not None else "")


def top_items(items: Counter[str], limit: int = 3, suffix: str = "个") -> str:
    if not items:
        return ""
    parts = []
    for name, count in items.most_common(limit):
        if not name:
            continue
        parts.append(f"{name}（{count}{suffix}）")
    return "、".join(parts)


def top_names(items: Counter[str], limit: int = 3) -> str:
    return "、".join([name for name, _ in items.most_common(limit) if name])


def format_many(prefix: str, count: int, detail: str | None = None) -> str:
    if detail:
        return f"{prefix}{count}个，{detail}"
    return f"{prefix}{count}个"


SUGGESTION_TEXTS = {
    "normal": "加大公共区域、小区（村）内的清扫保洁力度和垃圾清理频次，加强绿地、路面管护，健全堆物堆料长效治理机制，消除安全隐患，保持市容环境整洁有序；坚持疏堵结合，规划停车区域，引导群众规范停车。对问题高发的小区（村）要加大日常巡查管控力度，对照问题台账逐项整改，切实把精细化管理落到实处。针对问题数量同比上升的情况，梳理高发问题和高发区域，加密重点路段、小区（村）巡查频次，加强常态化管控，切实推动问题数量持续下降。",
    "waste": "持续开展垃圾分类宣传，提升居民自主分类意识，督促责任单位加强桶站满冒、周边不洁等突出问题的治理。进一步强化源头管控与动态巡查，对小卫星监测反馈及整改后易反弹的点位，定期组织“回头看”，巩固整改成效，严防问题反弹。",
    "furniture": "加强城市家具日常管护，健全基础台账，压紧压实各产权单位主体责任。针对市级检查中发现的破损问题，立行立改，举一反三，全面提升常态化管理水平。",
    "backstreet": "请加大日常巡查力度，重点聚焦条均问题突出的街巷，逐一排查问题，做到早发现、早处置、早整改，持续压降背街小巷条均问题数，不断提升背街小巷治理精细化水平。",
    "grid": "请对照网格倒查反馈清单，全面自查自纠，梳理漏报原因及薄弱环节；同时组织网格员开展专项培训，完善考核机制，将上报准确率、及时性与绩效挂钩，强化日常巡查管控，切实提升市区两级抽查问题的上报率和处置率。",
}

