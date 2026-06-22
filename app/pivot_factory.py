"""Creation and reuse of Power BI/OLAP PivotTables through Excel COM."""

from __future__ import annotations

import hashlib
import json
import logging
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    from filter_engine import FilterEngine, normalize_filters
except ImportError:
    from filter_engine import FilterEngine, normalize_filters

logger = logging.getLogger("pivot_factory")

XL_HIDDEN = 0
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_DATA_FIELD = 4


def _measure_descriptor(measure: Any) -> Dict[str, Any]:
    if isinstance(measure, dict):
        return dict(measure)
    return {
        "field_type": "named_measure",
        "measure_name": str(measure or "").strip(),
        "canonical_reference": str(measure or "").strip(),
    }


def _measure_signature(measure: Any) -> str:
    desc = _measure_descriptor(measure)
    return "|".join(
        [
            str(desc.get("field_type") or ""),
            str(desc.get("aggregation") or ""),
            str(desc.get("table_name") or ""),
            str(desc.get("column_name") or ""),
            str(desc.get("measure_name") or ""),
            str(desc.get("canonical_reference") or ""),
        ]
    ).casefold()


class PivotFactory:
    """Build connected PivotTables from a reusable OLAP PivotCache."""

    def __init__(
        self,
        *,
        pivot_cache: Any = None,
        template_pivot_table: Any = None,
        workbook_connection: Any = None,
    ) -> None:
        self.pivots: Dict[str, Dict[str, Any]] = {}
        self._olap_pivot_cache = pivot_cache
        self._template_pivot_table = template_pivot_table
        self._workbook_connection = workbook_connection
        self.created_count = 0
        self.failed_count = 0
        self.errors: List[Dict[str, Any]] = []
        self.sheet_allocation_rows: Dict[str, int] = {
            "_PBI_KPIs": 1,
            "_PBI_Charts": 1,
            "_PBI_Tables": 1,
            "_PBI_Filters": 1,
        }
        self.pivot_counter = 0
        self.filter_engine = FilterEngine()

    def generate_signature(
        self,
        rows: List[str],
        columns: List[str],
        measures: List[Any],
        legend: List[str],
        filters: Optional[Iterable[Any]] = None,
        sort_order: str = "",
        top_n: Optional[int] = None,
        unique_key: str = "",
    ) -> str:
        payload = {
            "rows": sorted(str(v).casefold() for v in rows if v),
            "columns": sorted(str(v).casefold() for v in columns if v),
            "measures": sorted(_measure_signature(v) for v in measures if v),
            "legend": sorted(str(v).casefold() for v in legend if v),
            "filters": normalize_filters(filters or []),
            "sort_order": str(sort_order or "").casefold(),
            "top_n": top_n,
            "unique_key": str(unique_key or "").casefold(),
        }
        return hashlib.sha256(
            json.dumps(payload, sort_keys=True, default=str).encode("utf-8")
        ).hexdigest()

    def allocate_helper_cell(
        self, sheet_name: str, estimated_rows: int = 40
    ) -> Tuple[int, int]:
        current_row = self.sheet_allocation_rows.get(sheet_name, 1)
        self.sheet_allocation_rows[sheet_name] = (
            current_row + max(estimated_rows, 25) + 5
        )
        return current_row, 1

    @staticmethod
    def _get_or_create_sheet(workbook: Any, sheet_name: str) -> Any:
        import time

        max_retries = 30
        delay = 0.5

        # Try to find the sheet first, with retries if Excel is busy
        for attempt in range(max_retries):
            try:
                return workbook.Worksheets(sheet_name)
            except Exception as e:
                err_str = str(e)
                # If Excel is busy, wait and retry
                if any(
                    x in err_str
                    for x in [
                        "-2146777998",
                        "800ac472",
                        "-2147418111",
                        "80010001",
                        "-2147417846",
                        "8001010a",
                    ]
                ):
                    if attempt < max_retries - 1:
                        time.sleep(delay)
                        continue
                # If it's a "sheet not found" error, break and create it
                break

        # If we reached here, the sheet was not found. Let's create it.
        for attempt in range(max_retries):
            try:
                sheet = workbook.Worksheets.Add()
                sheet.Name = sheet_name
                sheet.Visible = XL_HIDDEN
                return sheet
            except Exception as e:
                err_str = str(e)
                if (
                    any(
                        x in err_str
                        for x in [
                            "-2146777998",
                            "800ac472",
                            "-2147418111",
                            "80010001",
                            "-2147417846",
                            "8001010a",
                        ]
                    )
                    and attempt < max_retries - 1
                ):
                    time.sleep(delay)
                    continue
                raise

    @staticmethod
    def _find_template_pivot(workbook: Any) -> Tuple[Any, Any]:
        preferred: List[Tuple[Any, Any]] = []
        fallback: List[Tuple[Any, Any]] = []
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            sheet = workbook.Worksheets(index)
            try:
                count = int(sheet.PivotTables().Count)
            except Exception:
                continue
            for p_index in range(1, count + 1):
                try:
                    pivot = sheet.PivotTables(p_index)
                    if not bool(pivot.PivotCache().OLAP):
                        continue
                    pair = (sheet, pivot)
                    if (
                        "template" in str(sheet.Name).casefold()
                        or "template" in str(pivot.Name).casefold()
                    ):
                        preferred.append(pair)
                    else:
                        fallback.append(pair)
                except Exception:
                    continue
        candidates = preferred + fallback
        if not candidates:
            raise RuntimeError(
                "No connected OLAP template PivotTable exists. Create one Analyze-in-Excel PivotTable in the template workbook."
            )
        return candidates[0]

    def _find_olap_pivot_cache(self, workbook: Any) -> Any:
        """Return the live session's OLAP PivotCache without requiring a template PivotTable."""
        if self._olap_pivot_cache is not None:
            try:
                if bool(self._olap_pivot_cache.OLAP):
                    return self._olap_pivot_cache
            except Exception:
                self._olap_pivot_cache = None

        try:
            cache_count = int(workbook.PivotCaches().Count)
        except Exception:
            cache_count = 0

        for index in range(1, cache_count + 1):
            try:
                cache = workbook.PivotCaches().Item(index)
                if bool(cache.OLAP):
                    self._olap_pivot_cache = cache
                    return cache
            except Exception:
                continue

        if self._template_pivot_table is not None:
            try:
                cache = self._template_pivot_table.PivotCache()
                if bool(cache.OLAP):
                    self._olap_pivot_cache = cache
                    return cache
            except Exception:
                pass

        raise RuntimeError(
            "No connected OLAP PivotCache is available in the current live session. "
            "Insert an empty PivotTable from the Power BI semantic model before continuing."
        )

    @staticmethod
    def _safe_pivot_name(value: str) -> str:
        cleaned = re.sub(r"[^A-Za-z0-9_]", "_", str(value or "Pivot"))
        if not cleaned or cleaned[0].isdigit():
            cleaned = f"PVT_{cleaned}"
        return cleaned[:250]

    @staticmethod
    def _clear_pivot_layout(pivot_table: Any) -> None:
        try:
            pivot_table.ManualUpdate = True
        except Exception:
            pass
        try:
            cube_fields = pivot_table.CubeFields
            for idx in range(1, int(cube_fields.Count) + 1):
                try:
                    cube_fields.Item(idx).Orientation = XL_HIDDEN
                except Exception:
                    continue
        except Exception:
            pass
        try:
            pivot_table.ClearAllFilters()
        except Exception:
            pass

    @staticmethod
    def _set_cube_orientation(
        pivot_table: Any, field_name: str, orientation: int
    ) -> None:
        pivot_table.CubeFields(field_name).Orientation = orientation

    @staticmethod
    def _add_named_measure(pivot_table: Any, field_name: str) -> None:
        cube_field = pivot_table.CubeFields(field_name)
        try:
            cube_field.Orientation = XL_DATA_FIELD
        except Exception:
            pivot_table.AddDataField(cube_field)

    @staticmethod
    def _add_implicit_measure(
        pivot_table: Any, field_name: str, aggregation: str, caption: str
    ) -> None:
        """Attempt OLAP implicit aggregation without inventing a named measure."""
        cube_field = pivot_table.CubeFields(field_name)
        try:
            data_field = pivot_table.AddDataField(cube_field, caption)
        except Exception as exc:
            raise RuntimeError(
                f"Excel OLAP could not aggregate source column '{field_name}' as {aggregation}. Add a named Power BI measure for production use. Details: {exc}"
            ) from exc

        # Excel constants: xlSum=-4157, xlAverage=-4106, xlCount=-4112, xlMin=-4139, xlMax=-4136.
        function_map = {
            "SUM": -4157,
            "AVERAGE": -4106,
            "AVG": -4106,
            "COUNT": -4112,
            "COUNTA": -4112,
            "MIN": -4139,
            "MAX": -4136,
        }
        function_value = function_map.get(str(aggregation or "SUM").upper())
        if function_value is not None:
            try:
                data_field.Function = function_value
            except Exception:
                # Many OLAP providers do not permit changing Function. The field is still valid if AddDataField succeeded.
                pass

    @staticmethod
    def _external_destination(sheet: Any, row: int, col: int) -> str:
        try:
            return str(sheet.Cells(row, col).Address(True, True, 1, True))
        except Exception:
            return f"'{sheet.Name}'!R{row}C{col}"

    def _create_direct(
        self,
        pivot_cache: Any,
        helper_sheet: Any,
        row: int,
        col: int,
        pivot_name: str,
    ) -> Tuple[Any, str]:
        destination = self._external_destination(helper_sheet, row, col)
        helper_sheet.Activate()
        try:
            pivot = pivot_cache.CreatePivotTable(destination, pivot_name)
        except Exception:
            pivot = pivot_cache.CreatePivotTable(
                TableDestination=destination, TableName=pivot_name
            )
        return pivot, str(helper_sheet.Name)

    def _clone_template_pivot(
        self,
        workbook: Any,
        pivot_name: str,
    ) -> Tuple[Any, str]:
        """Copy the one user-created connected PivotTable inside the same workbook.

        Important: Worksheet.Copy(Before, After) is called with positional COM
        arguments.  With dynamic pywin32 dispatch, using Copy(After=sheet) can
        silently ignore the named argument and create a new workbook instead.
        """
        import time

        template_pivot = self._template_pivot_table
        if template_pivot is None:
            _, template_pivot = self._find_template_pivot(workbook)
            self._template_pivot_table = template_pivot

        try:
            template_sheet = template_pivot.Parent
            original_name = str(template_sheet.Name)
            original_parent = template_sheet.Parent
        except Exception as exc:
            raise RuntimeError(
                "The user-created connected PivotTable is no longer available."
            ) from exc

        # Snapshot the current workbook sheets before copying.
        before_names = set()
        before_count = int(workbook.Worksheets.Count)
        for index in range(1, before_count + 1):
            try:
                before_names.add(str(workbook.Worksheets(index).Name).casefold())
            except Exception:
                continue

        try:
            last_sheet = workbook.Worksheets(before_count)
            template_sheet.Activate()

            # COM signature: Copy(Before, After)
            # Use positional arguments so pywin32 keeps the copy in this workbook.
            template_sheet.Copy(None, last_sheet)
        except Exception as exc:
            raise RuntimeError(
                f"Excel could not copy the connected PivotTable sheet: {exc}"
            ) from exc

        cloned_sheet = None

        for _ in range(60):
            try:
                current_count = int(workbook.Worksheets.Count)
            except Exception:
                current_count = before_count

            # Preferred detection: find the new worksheet name in the same workbook.
            if current_count > before_count:
                for index in range(1, current_count + 1):
                    try:
                        candidate = workbook.Worksheets(index)
                        candidate_name = str(candidate.Name)
                        if candidate_name.casefold() not in before_names:
                            cloned_sheet = candidate
                            break
                    except Exception:
                        continue

            if cloned_sheet is not None:
                break

            # Secondary detection: Excel usually activates the copied sheet.
            try:
                active_sheet = workbook.Application.ActiveSheet
                active_parent = active_sheet.Parent
                active_name = str(active_sheet.Name)

                same_workbook = False
                try:
                    same_workbook = (
                        str(active_parent.FullName).casefold()
                        == str(workbook.FullName).casefold()
                    )
                except Exception:
                    same_workbook = active_parent is original_parent

                if (
                    same_workbook
                    and active_name.casefold() not in before_names
                ):
                    cloned_sheet = active_sheet
                    break
            except Exception:
                pass

            try:
                import pythoncom  # type: ignore[import]
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

            time.sleep(0.10)

        if cloned_sheet is None:
            try:
                active_book = workbook.Application.ActiveWorkbook
                active_book_name = str(active_book.Name)
                workbook_name = str(workbook.Name)

                if active_book_name.casefold() != workbook_name.casefold():
                    # Clean up a workbook accidentally created by Worksheet.Copy.
                    try:
                        active_book.Close(False)
                    except Exception:
                        pass

                    raise RuntimeError(
                        "Excel copied the PivotTable into a new workbook instead "
                        "of the live session workbook."
                    )
            except RuntimeError:
                raise
            except Exception:
                pass

            raise RuntimeError(
                "Excel did not add the copied PivotTable sheet to the live workbook."
            )

        base = f"_PBI_PVT_{self.pivot_counter:03d}"
        candidate_name = base[:31]

        existing_names = set()
        try:
            count = int(workbook.Worksheets.Count)
            for index in range(1, count + 1):
                sheet = workbook.Worksheets(index)
                name = str(sheet.Name)
                if name.casefold() != str(cloned_sheet.Name).casefold():
                    existing_names.add(name.casefold())
        except Exception:
            pass

        suffix = 2
        while candidate_name.casefold() in existing_names:
            tail = f"_{suffix}"
            candidate_name = f"{base[:31-len(tail)]}{tail}"
            suffix += 1

        try:
            cloned_sheet.Name = candidate_name
        except Exception:
            candidate_name = str(cloned_sheet.Name)

        # Keep visible until its fields are configured.
        try:
            cloned_sheet.Visible = -1
        except Exception:
            pass

        pivot_count = 0
        for _ in range(40):
            try:
                pivot_count = int(cloned_sheet.PivotTables().Count)
                if pivot_count > 0:
                    break
            except Exception:
                pass

            try:
                import pythoncom  # type: ignore[import]
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

            time.sleep(0.10)

        if pivot_count < 1:
            raise RuntimeError(
                "The copied worksheet does not contain the connected PivotTable."
            )

        cloned_pivot = cloned_sheet.PivotTables(1)

        try:
            if not bool(cloned_pivot.PivotCache().OLAP):
                raise RuntimeError(
                    "The copied PivotTable lost the Power BI OLAP connection."
                )
        except RuntimeError:
            raise
        except Exception as exc:
            raise RuntimeError(
                "The copied PivotTable connection could not be verified."
            ) from exc

        try:
            cloned_pivot.Name = pivot_name
        except Exception:
            pass

        return cloned_pivot, str(cloned_sheet.Name)

    def _configure_fields(
        self,
        pivot_table: Any,
        rows: List[str],
        columns: List[str],
        measures: List[Any],
        legend: List[str],
        field_mapper: Any,
    ) -> List[str]:
        warnings: List[str] = []
        self._clear_pivot_layout(pivot_table)

        for field in rows:
            mapping = field_mapper.map_field(field, "dimension")
            if mapping.get("status") != "mapped":
                warnings.append(
                    f"Row field could not be mapped: {field} ({mapping.get('status')})"
                )
                continue
            try:
                self._set_cube_orientation(
                    pivot_table, mapping["excel_olap_field"], XL_ROW_FIELD
                )
            except Exception as exc:
                warnings.append(f"Failed to add row field '{field}': {exc}")

        for field in [*columns, *legend]:
            mapping = field_mapper.map_field(field, "dimension")
            if mapping.get("status") != "mapped":
                warnings.append(
                    f"Column/legend field could not be mapped: {field} ({mapping.get('status')})"
                )
                continue
            try:
                self._set_cube_orientation(
                    pivot_table, mapping["excel_olap_field"], XL_COLUMN_FIELD
                )
            except Exception as exc:
                warnings.append(f"Failed to add column/legend field '{field}': {exc}")

        valid_measure_count = 0
        for raw_measure in measures:
            desc = _measure_descriptor(raw_measure)
            mapping = field_mapper.map_field(desc, "measure")
            semantic_type = str(
                desc.get("field_type") or mapping.get("field_type") or ""
            )
            try:
                if (
                    semantic_type == "named_measure"
                    and mapping.get("status") == "mapped"
                ):
                    self._add_named_measure(pivot_table, mapping["excel_olap_field"])
                    valid_measure_count += 1
                elif (
                    semantic_type == "implicit_measure"
                    and mapping.get("status") == "requires_pivot_aggregation"
                ):
                    caption = str(
                        desc.get("display_name")
                        or f"{desc.get('aggregation', 'Sum')} of {desc.get('column_name', 'Value')}"
                    )
                    self._add_implicit_measure(
                        pivot_table,
                        mapping["excel_olap_field"],
                        str(desc.get("aggregation") or "SUM"),
                        caption,
                    )
                    valid_measure_count += 1
                else:
                    warnings.append(
                        f"Measure could not be mapped safely: {desc.get('measure_name') or desc.get('canonical_reference') or desc.get('raw_reference')} ({mapping.get('status')})"
                    )
            except Exception as exc:
                warnings.append(str(exc))

        if measures and valid_measure_count == 0:
            raise RuntimeError(
                "No requested measure could be added to the connected PivotTable."
            )

        try:
            pivot_table.ManualUpdate = False
        except Exception:
            pass
        return warnings

    def get_or_create_pivot(
        self,
        excel_app: Any,
        workbook: Any,
        connection_name: str,
        rows: List[str],
        columns: List[str],
        measures: List[Any],
        legend: List[str],
        visual_type: str,
        field_mapper: Any,
        filters: Optional[Iterable[Any]] = None,
        sort_order: str = "",
        top_n: Optional[int] = None,
        unique_key: str = "",
    ) -> Dict[str, Any]:
        signature = self.generate_signature(
            rows,
            columns,
            measures,
            legend,
            filters,
            sort_order,
            top_n,
            unique_key,
        )
        if signature in self.pivots:
            return {**self.pivots[signature], "reused": True}

        self.pivot_counter += 1
        pivot_name = self._safe_pivot_name(f"PVT_PBIEX_{self.pivot_counter:03d}")
        visual_type_lower = str(visual_type or "").casefold()
        if visual_type_lower in {"table", "matrix"}:
            sheet_name = "_PBI_Tables"
        elif visual_type_lower in {"card", "kpi"}:
            sheet_name = "_PBI_KPIs"
        elif "slicer" in visual_type_lower or "filter" in visual_type_lower:
            sheet_name = "_PBI_Filters"
        else:
            sheet_name = "_PBI_Charts"

        helper_sheet = self._get_or_create_sheet(workbook, sheet_name)
        row, col = self.allocate_helper_cell(sheet_name)
        warnings: List[str] = []
        creation_mode = "direct"
        actual_sheet_name = sheet_name
        pivot_table = None

        try:
            pivot_cache = self._find_olap_pivot_cache(workbook)

            # Power BI Analyze-in-Excel PivotCaches commonly reject
            # CreatePivotTable.  The reliable workflow is:
            # 1) user inserts one connected PivotTable;
            # 2) pywin32 copies it for every visual/KPI;
            # 3) each copy is reconfigured with semantic-model CubeFields.
            creation_mode = "template_clone"
            pivot_table, actual_sheet_name = self._clone_template_pivot(
                workbook, pivot_name
            )

            warnings.extend(
                self._configure_fields(
                    pivot_table, rows, columns, measures, legend, field_mapper
                )
            )
            warnings.extend(
                self.filter_engine.apply_filters_to_pivot(
                    pivot_table, filters or [], field_mapper
                )
            )
            try:
                pivot_table.RefreshTable()
            except Exception as exc:
                warnings.append(f"PivotTable refresh warning: {exc}")

            try:
                pivot_table.Parent.Visible = XL_HIDDEN
            except Exception as exc:
                warnings.append(f"Could not hide PivotTable helper sheet: {exc}")

            info = {
                "pivot_name": str(pivot_table.Name),
                "sheet_name": actual_sheet_name,
                "row": row,
                "col": col,
                "range_address": f"'{actual_sheet_name}'!R{row}C{col}",
                "signature": signature,
                "reused": False,
                "warnings": warnings,
                "pivot_table": pivot_table,
                "creation_mode": creation_mode,
                "connection_name": connection_name,
            }
            self.pivots[signature] = info
            self.created_count += 1
            return info
        except Exception as exc:
            logger.exception("Failed to build connected PivotTable %s", pivot_name)
            self.failed_count += 1
            self.errors.append(
                {
                    "pivot_name": pivot_name,
                    "signature": signature,
                    "error": str(exc),
                }
            )
            return {
                "pivot_name": pivot_name,
                "sheet_name": actual_sheet_name,
                "row": row,
                "col": col,
                "range_address": f"'{actual_sheet_name}'!R{row}C{col}",
                "signature": signature,
                "reused": False,
                "warnings": warnings,
                "error": str(exc),
                "creation_mode": creation_mode,
                "connection_name": connection_name,
            }

    def create_filter_source_pivot(
        self,
        excel_app: Any,
        workbook: Any,
        connection_name: str,
        field: str,
        field_mapper: Any,
    ) -> Dict[str, Any]:
        return self.get_or_create_pivot(
            excel_app=excel_app,
            workbook=workbook,
            connection_name=connection_name,
            rows=[field],
            columns=[],
            measures=[],
            legend=[],
            visual_type="slicer",
            field_mapper=field_mapper,
            filters=[],
            unique_key=f"filter:{field}",
        )


__all__ = ["PivotFactory"]
