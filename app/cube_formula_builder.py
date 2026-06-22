"""Production-grade builders for Excel CUBE formulas.

This module generates validated Excel CUBE formulas for workbooks connected to
Power BI semantic models or other OLAP sources.

The DAX logic remains inside the semantic model. Excel formulas generated here
reference existing model measures, OLAP members, slicers, named ranges, or cells.
"""

from __future__ import annotations

import math
import re
from numbers import Real
from typing import Any, Dict, Iterable, List, Optional, Sequence, Union

FilterMember = Union[str, Dict[str, Any]]

_VALID_CUBESET_SORT_ORDERS = {0, 1, 2, 3, 4, 5, 6}
_MAX_EXCEL_FORMULA_LENGTH = 8192


class CubeFormulaError(ValueError):
    """Raised when a CUBE formula cannot be created safely."""


def _require_text(value: Any, field_name: str) -> str:
    """Return a trimmed string or raise a clear validation error."""
    text = str(value or "").strip()
    if not text:
        raise CubeFormulaError(f"{field_name} is required.")
    return text


def _excel_string(value: Any) -> str:
    """Return a safely quoted Excel string literal."""
    return '"' + str(value).replace('"', '""') + '"'


def _excel_literal(value: Any) -> str:
    """Return an Excel-safe literal without converting numbers into text."""
    if value is None:
        return '""'

    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"

    if isinstance(value, Real):
        numeric = float(value)

        if not math.isfinite(numeric):
            raise CubeFormulaError("Excel fallback values must be finite numbers.")

        if isinstance(value, int) or numeric.is_integer():
            return str(int(numeric))

        return format(numeric, ".15g")

    return _excel_string(value)


def _strip_formula_prefix(value: str) -> str:
    """Remove a leading equals sign from an Excel expression."""
    text = str(value or "").strip()
    return text[1:].strip() if text.startswith("=") else text


def _is_formula_reference(value: str) -> bool:
    """Return True when a value looks like an Excel name or reference."""
    text = _strip_formula_prefix(value)

    if not text:
        return False

    defined_name = r"[A-Za-z_\\][A-Za-z0-9_.\\]*"
    cell = r"\$?[A-Z]{1,3}\$?[1-9][0-9]*"
    cell_range = rf"{cell}(?::{cell})?"
    quoted_sheet = r"'(?:[^']|'')+'"
    unquoted_sheet = r"[A-Za-z_][A-Za-z0-9_ ]*"
    sheet = rf"(?:{quoted_sheet}|{unquoted_sheet})"

    if re.fullmatch(defined_name, text):
        return True

    if re.fullmatch(rf"{cell_range}#?", text):
        return True

    if re.fullmatch(rf"{sheet}!{cell_range}#?", text):
        return True

    return False


def _looks_like_olap_expression(value: str) -> bool:
    """Return True when text resembles an OLAP unique name or set expression."""
    text = str(value or "").strip()

    if not text:
        return False

    if text.startswith("{") and text.endswith("}"):
        return True

    return text.startswith("[") and "]" in text


def _normalize_measure_path(measure_olap_path: str) -> str:
    """Normalize a measure name to an OLAP measure unique name."""
    measure = _require_text(measure_olap_path, "measure_olap_path")

    if measure.casefold().startswith("[measures]."):
        return measure

    if measure.startswith("[") and measure.endswith("]"):
        return f"[Measures].{measure}"

    escaped = measure.replace("]", "]]")
    return f"[Measures].[{escaped}]"


def _build_filter_argument(member: FilterMember) -> Optional[str]:
    """Convert a filter descriptor into a valid CUBEVALUE argument."""
    if isinstance(member, dict):
        formula_reference = (
            member.get("formula_reference")
            or member.get("slicer_reference")
            or member.get("cell_reference")
            or member.get("named_range")
        )
        olap_member = (
            member.get("olap_member")
            or member.get("member")
            or member.get("member_unique_name")
            or member.get("set_expression")
        )

        if formula_reference:
            reference = _strip_formula_prefix(str(formula_reference))

            if not _is_formula_reference(reference):
                raise CubeFormulaError(
                    f"Invalid Excel formula reference: {formula_reference}"
                )

            return reference

        if olap_member:
            expression = _require_text(
                olap_member,
                "filter member OLAP path",
            )

            if not _looks_like_olap_expression(expression):
                raise CubeFormulaError(
                    "OLAP filter members must be unique names or set expressions: "
                    f"{expression}"
                )

            return _excel_string(expression)

        return None

    text = str(member or "").strip()
    if not text:
        return None

    expression = _strip_formula_prefix(text)

    if _is_formula_reference(expression):
        return expression

    if not _looks_like_olap_expression(text):
        raise CubeFormulaError(
            "Filter values must be an Excel reference, slicer name, named range, "
            f"or OLAP member/set expression: {text}"
        )

    return _excel_string(text)


def _wrap_iferror(formula: str, fallback: Any = "") -> str:
    """Wrap a formula in IFERROR while preserving numeric fallback types."""
    expression = _strip_formula_prefix(formula)
    return f"=IFERROR({expression}, {_excel_literal(fallback)})"


def _validate_formula_length(formula: str) -> str:
    """Reject formulas that exceed Excel's formula-length limit."""
    if len(formula) > _MAX_EXCEL_FORMULA_LENGTH:
        raise CubeFormulaError(
            f"Generated formula is too long for Excel: {len(formula)} characters."
        )
    return formula


def _finalize_formula(
    formula: str,
    *,
    use_iferror: bool,
    error_fallback: Any,
) -> str:
    """Apply optional IFERROR wrapping and validate final length."""
    final_formula = (
        _wrap_iferror(formula, error_fallback)
        if use_iferror
        else formula
    )
    return _validate_formula_length(final_formula)


def _normalize_rank(rank: Union[int, str]) -> str:
    """Normalize a CUBERANKEDMEMBER rank argument."""
    if isinstance(rank, bool):
        raise CubeFormulaError("rank must be a positive integer or cell reference.")

    if isinstance(rank, int):
        if rank < 1:
            raise CubeFormulaError("rank must be greater than or equal to 1.")
        return str(rank)

    expression = _strip_formula_prefix(str(rank or "").strip())

    if not expression:
        raise CubeFormulaError("rank is required.")

    if expression.isdigit():
        numeric_rank = int(expression)

        if numeric_rank < 1:
            raise CubeFormulaError("rank must be greater than or equal to 1.")

        return str(numeric_rank)

    if _is_formula_reference(expression):
        return expression

    raise CubeFormulaError(
        f"rank must be a positive integer or Excel reference: {rank}"
    )


def build_cube_value_formula(
    connection_name: str,
    measure_olap_path: str,
    filter_members: Optional[Iterable[FilterMember]] = None,
    *,
    use_iferror: bool = True,
    error_fallback: Any = "",
) -> str:
    """Build a CUBEVALUE formula for an existing semantic-model measure."""
    connection = _require_text(connection_name, "connection_name")
    measure = _normalize_measure_path(measure_olap_path)

    arguments: List[str] = [
        _excel_string(connection),
        _excel_string(measure),
    ]

    for member in filter_members or []:
        argument = _build_filter_argument(member)
        if argument:
            arguments.append(argument)

    formula = f'=CUBEVALUE({", ".join(arguments)})'

    return _finalize_formula(
        formula,
        use_iferror=use_iferror,
        error_fallback=error_fallback,
    )


def build_cube_member_formula(
    connection_name: str,
    member_olap_path: str,
    *,
    caption: Optional[str] = None,
    use_iferror: bool = True,
    error_fallback: Any = "",
) -> str:
    """Build a CUBEMEMBER formula for a dimension member."""
    connection = _require_text(connection_name, "connection_name")
    member = _require_text(member_olap_path, "member_olap_path")

    if not _looks_like_olap_expression(member):
        raise CubeFormulaError(
            f"member_olap_path must be an OLAP unique name: {member}"
        )

    arguments: List[str] = [
        _excel_string(connection),
        _excel_string(member),
    ]

    if caption is not None:
        arguments.append(_excel_string(caption))

    formula = f'=CUBEMEMBER({", ".join(arguments)})'

    return _finalize_formula(
        formula,
        use_iferror=use_iferror,
        error_fallback=error_fallback,
    )


def build_cube_set_formula(
    connection_name: str,
    set_expression: str,
    caption: str,
    *,
    sort_order: Optional[int] = None,
    sort_by: Optional[str] = None,
    use_iferror: bool = True,
    error_fallback: Any = "",
) -> str:
    """Build a CUBESET formula for a set of OLAP members."""
    connection = _require_text(connection_name, "connection_name")
    expression = _require_text(set_expression, "set_expression")
    set_caption = _require_text(caption, "caption")

    if not _looks_like_olap_expression(expression):
        raise CubeFormulaError(
            f"set_expression must be an OLAP set expression: {expression}"
        )

    arguments: List[str] = [
        _excel_string(connection),
        _excel_string(expression),
        _excel_string(set_caption),
    ]

    if sort_order is not None:
        normalized_sort_order = int(sort_order)

        if normalized_sort_order not in _VALID_CUBESET_SORT_ORDERS:
            raise CubeFormulaError(
                "sort_order must be an Excel CUBESET sort value from 0 through 6."
            )

        arguments.append(str(normalized_sort_order))

        if sort_by is not None:
            arguments.append(
                _excel_string(
                    _require_text(sort_by, "sort_by")
                )
            )

    elif sort_by is not None:
        raise CubeFormulaError(
            "sort_order is required when sort_by is provided."
        )

    formula = f'=CUBESET({", ".join(arguments)})'

    return _finalize_formula(
        formula,
        use_iferror=use_iferror,
        error_fallback=error_fallback,
    )


def build_cube_ranked_member_formula(
    connection_name: str,
    set_reference: str,
    rank: Union[int, str],
    *,
    caption: Optional[str] = None,
    use_iferror: bool = True,
    error_fallback: Any = "",
) -> str:
    """Build a CUBERANKEDMEMBER formula from a CUBESET reference."""
    connection = _require_text(connection_name, "connection_name")
    set_ref = _strip_formula_prefix(
        _require_text(set_reference, "set_reference")
    )

    if not _is_formula_reference(set_ref):
        raise CubeFormulaError(
            f"Invalid CUBESET formula reference: {set_reference}"
        )

    rank_expression = _normalize_rank(rank)

    arguments: List[str] = [
        _excel_string(connection),
        set_ref,
        rank_expression,
    ]

    if caption is not None:
        arguments.append(_excel_string(caption))

    formula = f'=CUBERANKEDMEMBER({", ".join(arguments)})'

    return _finalize_formula(
        formula,
        use_iferror=use_iferror,
        error_fallback=error_fallback,
    )


def build_cube_member_property_formula(
    connection_name: str,
    member_reference: str,
    property_name: str,
    *,
    use_iferror: bool = True,
    error_fallback: Any = "",
) -> str:
    """Build a CUBEMEMBERPROPERTY formula."""
    connection = _require_text(connection_name, "connection_name")
    member_ref = _strip_formula_prefix(
        _require_text(member_reference, "member_reference")
    )
    property_text = _require_text(property_name, "property_name")

    if not _is_formula_reference(member_ref):
        raise CubeFormulaError(
            f"Invalid CUBEMEMBER reference: {member_reference}"
        )

    formula = (
        f"=CUBEMEMBERPROPERTY("
        f"{_excel_string(connection)}, "
        f"{member_ref}, "
        f"{_excel_string(property_text)}"
        f")"
    )

    return _finalize_formula(
        formula,
        use_iferror=use_iferror,
        error_fallback=error_fallback,
    )


def build_cube_set_count_formula(
    set_reference: str,
    *,
    use_iferror: bool = True,
    error_fallback: Any = 0,
) -> str:
    """Build a CUBESETCOUNT formula."""
    set_ref = _strip_formula_prefix(
        _require_text(set_reference, "set_reference")
    )

    if not _is_formula_reference(set_ref):
        raise CubeFormulaError(
            f"Invalid CUBESET reference: {set_reference}"
        )

    formula = f"=CUBESETCOUNT({set_ref})"

    return _finalize_formula(
        formula,
        use_iferror=use_iferror,
        error_fallback=error_fallback,
    )


def build_measure_reference(measure_name: str) -> str:
    """Return a normalized OLAP measure path."""
    return _normalize_measure_path(measure_name)


__all__: Sequence[str] = (
    "CubeFormulaError",
    "build_cube_value_formula",
    "build_cube_member_formula",
    "build_cube_set_formula",
    "build_cube_ranked_member_formula",
    "build_cube_member_property_formula",
    "build_cube_set_count_formula",
    "build_measure_reference",
)
