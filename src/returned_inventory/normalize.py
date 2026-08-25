"""M04 数据标准化（对应 VBA modNormalize.bas）。

职责：
1. 把 excel_input 读取的 Raw* 原始数据转换为强类型 Normalized* 数据。
2. 保留字段是否合法的标记（line_no_valid / qty_valid / expiry_valid 等）。
3. 对每个非法字段生成 FieldNormalizeIssue，供 validate 输出异常明细。

注意：标准化阶段只记录"字段问题"，不直接生成 E01/E03/E04/E05 等错误码；
错误码由 validate 校验层统一生成。所有函数不接触 Excel 对象。
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from .models import (
    CELL_KIND_BLANK,
    CELL_KIND_EXCEL_DATE,
    CELL_KIND_TEXT,
    ISSUE_KIND_EMPTY,
    ISSUE_KIND_FORMAT_ERROR,
    ISSUE_KIND_RANGE_ERROR,
    QC_NG,
    QC_QC,
    QC_ZP,
    SOURCE_INVENTORY_TABLE,
    SOURCE_RETURN_TABLE,
    Config,
    FieldNormalizeIssue,
    NormalizedInventoryLine,
    NormalizedReturnLine,
    RawInventoryRow,
    RawReturnRow,
)

# -----------------------------------------------------------------------------
# 公开函数
# -----------------------------------------------------------------------------


def normalize_return_rows(
    raws: list[RawReturnRow], cfg: Config
) -> tuple[list[NormalizedReturnLine], list[FieldNormalizeIssue]]:
    """cfg 为与库存标准化保持一致的预留参数；当前退单字段标准化不读取配置值。"""
    del cfg
    issues: list[FieldNormalizeIssue] = []
    return [_normalize_return_line(raw, issues) for raw in raws], issues


def normalize_inventory_rows(
    raws: list[RawInventoryRow], cfg: Config
) -> tuple[list[NormalizedInventoryLine], list[FieldNormalizeIssue]]:
    issues: list[FieldNormalizeIssue] = []
    return [_normalize_inventory_line(raw, cfg, issues) for raw in raws], issues


# -----------------------------------------------------------------------------
# 单行标准化
# -----------------------------------------------------------------------------


def _normalize_return_line(
    raw: RawReturnRow, issues: list[FieldNormalizeIssue]
) -> NormalizedReturnLine:
    empty_fields: list[str] = []

    shipment_no, valid, kind = normalize_required_text(raw.shipment_no)
    if not valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_RETURN_TABLE, "物流单号", raw.shipment_no, kind, empty_fields)

    wms_order_no, valid, kind = normalize_required_text(raw.wms_order_no)
    if not valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_RETURN_TABLE, "WMS退单号", raw.wms_order_no, kind, empty_fields)

    sku, valid, kind = normalize_required_text(raw.sku)
    if not valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_RETURN_TABLE, "SKU", raw.sku, kind, empty_fields)

    line_no, line_no_valid, kind = normalize_line_no(raw.line_no)
    if not line_no_valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_RETURN_TABLE, "行号", raw.line_no, kind, empty_fields)

    qty, qty_valid, kind = normalize_qty(raw.qty)
    if not qty_valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_RETURN_TABLE, "数量", raw.qty, kind, empty_fields)

    return NormalizedReturnLine(
        excel_row_num=raw.excel_row_num,
        shipment_no=shipment_no,
        wms_order_no=wms_order_no,
        sku=sku,
        line_no=line_no,
        qty=qty,
        line_no_valid=line_no_valid,
        qty_valid=qty_valid,
        empty_fields=",".join(empty_fields),
    )


def _normalize_inventory_line(
    raw: RawInventoryRow, cfg: Config, issues: list[FieldNormalizeIssue]
) -> NormalizedInventoryLine:
    empty_fields: list[str] = []

    shipment_no, valid, kind = normalize_required_text(raw.shipment_no)
    if not valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_INVENTORY_TABLE, "物流单号", raw.shipment_no, kind, empty_fields)

    sku, valid, kind = normalize_required_text(raw.sku)
    if not valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_INVENTORY_TABLE, "SKU", raw.sku, kind, empty_fields)

    qc, qc_valid, kind = normalize_qc(raw.qc)
    if not qc_valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_INVENTORY_TABLE, "QC情况", raw.qc, kind, empty_fields)

    lot_no, valid, kind = normalize_lot_no(raw.lot_no, cfg)
    if not valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_INVENTORY_TABLE, "批号", raw.lot_no, kind, empty_fields)

    expiry, expiry_valid, kind = normalize_expiry(raw.expiry, raw.expiry_cell_kind)
    if not expiry_valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_INVENTORY_TABLE, "效期", raw.expiry, kind, empty_fields)

    qty, qty_valid, kind = normalize_qty(raw.qty)
    if not qty_valid:
        _add_field_issue(issues, raw.excel_row_num, SOURCE_INVENTORY_TABLE, "数量", raw.qty, kind, empty_fields)

    return NormalizedInventoryLine(
        excel_row_num=raw.excel_row_num,
        shipment_no=shipment_no,
        sku=sku,
        qc=qc,
        lot_no=lot_no,
        expiry=expiry,
        qty=qty,
        qc_valid=qc_valid,
        expiry_valid=expiry_valid,
        qty_valid=qty_valid,
        empty_fields=",".join(empty_fields),
    )


# -----------------------------------------------------------------------------
# 字段标准化（纯函数：返回 (结果, 是否合法, 问题类型)）
# -----------------------------------------------------------------------------


def normalize_required_text(raw_value: Any) -> tuple[str, bool, str]:
    result = _variant_to_text(raw_value).strip()
    if not result:
        return result, False, ISSUE_KIND_EMPTY
    return result, True, ""


def normalize_line_no(raw_value: Any) -> tuple[str, bool, str]:
    """行号：恰好 5 位且全为数字字符才合法；不自动补零（需求 §2.1）。"""
    result = _variant_to_text(raw_value).strip()
    if not result:
        return result, False, ISSUE_KIND_EMPTY
    if len(result) == 5 and _is_all_digits(result):
        return result, True, ""
    return result, False, ISSUE_KIND_FORMAT_ERROR


def normalize_qc(raw_value: Any) -> tuple[str, bool, str]:
    """QC：strip → 大写 → 校验，三步顺序不可颠倒（需求 §4.0）。"""
    result = _variant_to_text(raw_value).strip().upper()
    if not result:
        return result, False, ISSUE_KIND_EMPTY
    if result in (QC_ZP, QC_QC, QC_NG):
        return result, True, ""
    return result, False, ISSUE_KIND_FORMAT_ERROR


def normalize_lot_no(raw_value: Any, cfg: Config) -> tuple[str, bool, str]:
    """批号：strip 后默认统一大写（可由配置切换为大小写敏感）。"""
    result = _variant_to_text(raw_value).strip()
    if not result:
        return result, False, ISSUE_KIND_EMPTY
    if not cfg.lot_case_sensitive:
        result = result.upper()
    return result, True, ""


def normalize_qty(raw_value: Any) -> tuple[int, bool, str]:
    """数量：必须是正整数；非数字 → FormatError，负数/零/小数 → RangeError。"""
    text = _variant_to_text(raw_value).strip()
    if not text:
        return 0, False, ISSUE_KIND_EMPTY

    try:
        number = float(text)
    except ValueError:
        return 0, False, ISSUE_KIND_FORMAT_ERROR

    if number <= 0 or number != int(number):
        return 0, False, ISSUE_KIND_RANGE_ERROR

    return int(number), True, ""


def normalize_expiry(raw_value: Any, cell_kind: str) -> tuple[str, bool, str]:
    """效期三分支（需求 §2.2）：Excel 日期序列号 / 文本（先字符串级校验）/ 空。"""
    if cell_kind == CELL_KIND_EXCEL_DATE:
        if isinstance(raw_value, (datetime, date)):
            return _format_date(raw_value), True, ""
        return "", False, ISSUE_KIND_FORMAT_ERROR

    if cell_kind == CELL_KIND_TEXT:
        result, is_valid = normalize_text_date(_variant_to_text(raw_value).strip())
        return result, is_valid, "" if is_valid else ISSUE_KIND_FORMAT_ERROR

    if cell_kind == CELL_KIND_BLANK:
        return "", False, ISSUE_KIND_EMPTY

    return "", False, ISSUE_KIND_FORMAT_ERROR


# -----------------------------------------------------------------------------
# 日期文本校验
# -----------------------------------------------------------------------------


def normalize_text_date(text_value: str) -> tuple[str, bool]:
    """YYYY/MM/DD 或 YYYY-MM-DD 字符串级校验（含闰年与各月天数上限）。"""
    if not text_value:
        return "", False

    if "/" in text_value:
        separator = "/"
    elif "-" in text_value:
        separator = "-"
    else:
        return "", False

    parts = text_value.split(separator)
    if len(parts) != 3:
        return "", False

    if len(parts[0]) != 4 or len(parts[1]) != 2 or len(parts[2]) != 2:
        return "", False

    if not all(_is_all_digits(p) for p in parts):
        return "", False

    year_value, month_value, day_value = (int(p) for p in parts)

    if year_value < 1900 or year_value > 2999:
        return "", False
    if month_value < 1 or month_value > 12:
        return "", False
    if day_value < 1 or day_value > _days_in_month(year_value, month_value):
        return "", False

    return f"{year_value:04d}/{month_value:02d}/{day_value:02d}", True


def _days_in_month(year_value: int, month_value: int) -> int:
    if month_value in (1, 3, 5, 7, 8, 10, 12):
        return 31
    if month_value in (4, 6, 9, 11):
        return 30
    return 29 if _is_leap_year(year_value) else 28


def _is_leap_year(year_value: int) -> bool:
    return (year_value % 4 == 0 and year_value % 100 != 0) or (year_value % 400 == 0)


def _format_date(value: datetime | date) -> str:
    return f"{value.year:04d}/{value.month:02d}/{value.day:02d}"


# -----------------------------------------------------------------------------
# Issue 记录工具
# -----------------------------------------------------------------------------


def _add_field_issue(
    issues: list[FieldNormalizeIssue],
    excel_row_num: int,
    source_table: str,
    field_name: str,
    raw_value: Any,
    issue_kind: str,
    empty_fields: list[str],
) -> None:
    if issue_kind == ISSUE_KIND_EMPTY:
        empty_fields.append(field_name)
    issues.append(
        FieldNormalizeIssue(
            excel_row_num=excel_row_num,
            source_table=source_table,
            field_name=field_name,
            raw_value=_variant_to_text(raw_value),
            issue_kind=issue_kind,
        )
    )


# -----------------------------------------------------------------------------
# 通用工具
# -----------------------------------------------------------------------------


def _variant_to_text(raw_value: Any) -> str:
    """对应 VBA VariantToText：空值/错误 → 空串；其余 → CStr 等价物。"""
    if raw_value is None:
        return ""
    if isinstance(raw_value, bool):
        return str(raw_value)
    if isinstance(raw_value, float) and raw_value.is_integer():
        # VBA CStr(1.0) = "1"，避免 Python str(1.0) = "1.0" 的差异
        return str(int(raw_value))
    return str(raw_value)


def _is_all_digits(text_value: str) -> bool:
    return bool(text_value) and all("0" <= ch <= "9" for ch in text_value)
