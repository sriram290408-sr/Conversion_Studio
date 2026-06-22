"""Classify DAX measures for Excel rendering.

For live Power BI / OLAP-connected workbooks, every named DAX measure should
normally remain in the semantic model and be referenced through CUBEVALUE.

This module still detects simple aggregations for diagnostics, but it does not
recommend converting them into standalone Excel formulas when live mode is used.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Sequence

SEMANTIC_MODEL_KEYWORDS: Sequence[str] = (
    "CALCULATE",
    "CALCULATETABLE",
    "ALL",
    "ALLEXCEPT",
    "ALLNOBLANKROW",
    "ALLSELECTED",
    "REMOVEFILTERS",
    "KEEPFILTERS",
    "FILTER",
    "TREATAS",
    "USERELATIONSHIP",
    "CROSSFILTER",
    "RELATED",
    "RELATEDTABLE",
    "SAMEPERIODLASTYEAR",
    "TOTALYTD",
    "TOTALQTD",
    "TOTALMTD",
    "DATESYTD",
    "DATESQTD",
    "DATESMTD",
    "DATEADD",
    "PARALLELPERIOD",
    "PREVIOUSDAY",
    "PREVIOUSMONTH",
    "PREVIOUSQUARTER",
    "PREVIOUSYEAR",
    "NEXTDAY",
    "NEXTMONTH",
    "NEXTQUARTER",
    "NEXTYEAR",
    "CLOSINGBALANCEMONTH",
    "CLOSINGBALANCEQUARTER",
    "CLOSINGBALANCEYEAR",
    "OPENINGBALANCEMONTH",
    "OPENINGBALANCEQUARTER",
    "OPENINGBALANCEYEAR",
    "LASTDATE",
    "FIRSTDATE",
    "STARTOFMONTH",
    "STARTOFQUARTER",
    "STARTOFYEAR",
    "ENDOFMONTH",
    "ENDOFQUARTER",
    "ENDOFYEAR",
    "HASONEVALUE",
    "ISFILTERED",
    "ISCROSSFILTERED",
    "SELECTEDVALUE",
    "VALUES",
    "SUMX",
    "AVERAGEX",
    "MINX",
    "MAXX",
    "COUNTX",
    "RANKX",
)

_SIMPLE_AGGREGATION_PATTERNS: Sequence[str] = (
    r"^\s*SUM\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*AVERAGE\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*MIN\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*MAX\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*COUNT\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*COUNTA\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*DISTINCTCOUNT\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\[[^\]]+\]\s*\)\s*$",
    r"^\s*COUNTROWS\s*\(\s*(?:'[^']+'|[A-Za-z0-9_ ]+)\s*\)\s*$",
)


def _normalize_formula(dax_formula: Any) -> str:
    return str(dax_formula or "").strip()


def _matched_semantic_keywords(formula_upper: str) -> List[str]:
    matches: List[str] = []

    for keyword in SEMANTIC_MODEL_KEYWORDS:
        if re.search(
            rf"\b{re.escape(keyword)}\b",
            formula_upper,
        ):
            matches.append(keyword)

    return matches


def _is_simple_aggregation(formula: str) -> bool:
    return any(
        re.fullmatch(pattern, formula, flags=re.IGNORECASE)
        for pattern in _SIMPLE_AGGREGATION_PATTERNS
    )


def _max_parenthesis_depth(formula: str) -> int:
    depth = 0
    maximum = 0

    for character in formula:
        if character == "(":
            depth += 1
            maximum = max(maximum, depth)
        elif character == ")":
            depth = max(0, depth - 1)

    return maximum


def classify_dax_measure(
    measure_name: str,
    dax_formula: str,
    *,
    live_semantic_model: bool = True,
) -> Dict[str, Any]:
    """Classify a DAX measure for Excel rendering.

    Parameters
    ----------
    measure_name:
        Existing Power BI semantic-model measure name.

    dax_formula:
        Original DAX expression.

    live_semantic_model:
        When True, the measure is always rendered through CUBEVALUE because the
        semantic model is the source of truth. Simple aggregations are detected
        only for diagnostics.

        When False, simple aggregations may be marked as candidates for a normal
        Excel formula fallback.

    Returns
    -------
    dict
        Classification and recommended rendering metadata.
    """
    name = str(measure_name or "").strip()
    formula = _normalize_formula(dax_formula)
    formula_upper = formula.upper()

    if not name:
        return {
            "measure_name": "",
            "classification": "invalid_measure",
            "recommended_binding": "unmapped",
            "reason": "Measure name is empty.",
            "matched_keywords": [],
            "is_simple_aggregation": False,
            "parenthesis_depth": 0,
            "cube_measure_path": "",
        }

    cube_measure_path = (
        name
        if name.lower().startswith("[measures].")
        else f"[Measures].[{name.replace(']', ']]')}]"
    )

    if not formula:
        return {
            "measure_name": name,
            "classification": "semantic_model_measure",
            "recommended_binding": "cube_formula",
            "reason": (
                "No DAX expression was available, so the existing semantic-model "
                "measure must be referenced directly."
            ),
            "matched_keywords": [],
            "is_simple_aggregation": False,
            "parenthesis_depth": 0,
            "cube_measure_path": cube_measure_path,
        }

    matched_keywords = _matched_semantic_keywords(formula_upper)
    is_simple = _is_simple_aggregation(formula)
    depth = _max_parenthesis_depth(formula)

    if live_semantic_model:
        if matched_keywords:
            reason = "Uses semantic-model context functions: " + ", ".join(
                matched_keywords
            )
        elif is_simple:
            reason = (
                "Simple aggregation detected, but live mode preserves the "
                "Power BI measure as the source of truth."
            )
        else:
            reason = (
                "Named DAX measures in a live-connected workbook should be "
                "evaluated by the Power BI semantic model."
            )

        return {
            "measure_name": name,
            "classification": "semantic_model_measure",
            "recommended_binding": "cube_formula",
            "reason": reason,
            "matched_keywords": matched_keywords,
            "is_simple_aggregation": is_simple,
            "parenthesis_depth": depth,
            "cube_measure_path": cube_measure_path,
        }

    if is_simple and not matched_keywords:
        return {
            "measure_name": name,
            "classification": "excel_formula_candidate",
            "recommended_binding": "excel_formula",
            "reason": (
                "Simple standalone aggregation detected and live semantic-model "
                "mode is disabled."
            ),
            "matched_keywords": [],
            "is_simple_aggregation": True,
            "parenthesis_depth": depth,
            "cube_measure_path": cube_measure_path,
        }

    return {
        "measure_name": name,
        "classification": "semantic_model_required",
        "recommended_binding": "cube_formula",
        "reason": ("Complex DAX or filter-context logic requires the semantic model."),
        "matched_keywords": matched_keywords,
        "is_simple_aggregation": is_simple,
        "parenthesis_depth": depth,
        "cube_measure_path": cube_measure_path,
    }


__all__ = [
    "SEMANTIC_MODEL_KEYWORDS",
    "classify_dax_measure",
]
