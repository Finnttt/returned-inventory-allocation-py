"""M03 数据加载（对应 VBA modExcelInput.bas）。

职责：
1. 从 Excel 工作表读取原始数据到 RawReturnRow / RawInventoryRow。
2. 保留单元格原始值，不做 strip、不补零、不转业务类型。
3. 对效期列额外记录 expiry_cell_kind，供 normalize 选择正确的标准化路径。

已锁定输入表结构：
- 输入_退单表：物流单号 | WMS退单号 | SKU | 行号 | 数量
- 输入_质检库存表：物流单号 | SKU | QC情况 | 批号 | 效期 | 数量 | 备注
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from .models import (
    CELL_KIND_BLANK,
    CELL_KIND_EXCEL_DATE,
    CELL_KIND_OTHER,
    CELL_KIND_TEXT,
    RawInventoryRow,
    RawReturnRow,
)

# 退单表列号（1-based）
COL_RETURN_SHIPMENT_NO = 1
COL_RETURN_WMS_ORDER_NO = 2
COL_RETURN_SKU = 3
COL_RETURN_LINE_NO = 4
COL_RETURN_QTY = 5

# 质检库存表列号（1-based）
COL_INV_SHIPMENT_NO = 1
COL_INV_SKU = 2
COL_INV_QC = 3
COL_INV_LOT_NO = 4
COL_INV_EXPIRY = 5
COL_INV_QTY = 6

RETURN_HEADERS = ["物流单号", "WMS退单号", "SKU", "行号", "数量"]
INVENTORY_HEADERS = ["物流单号", "SKU", "QC情况", "批号", "效期", "数量"]


class InputError(Exception):
    """E12 结构异常：表头不符或 WMS退单号跨物流单号重复，命中即中止整次运行。"""


def read_return_orders(ws: Any) -> list[RawReturnRow]:
    """读取退单表。加载过程中同步检查 WMS退单号 全局唯一（E12-②）。"""
    _validate_header(ws, RETURN_HEADERS)

    result: list[RawReturnRow] = []
    wms_to_shipment: dict[str, str] = {}

    for row_index in range(2, _last_used_row(ws, len(RETURN_HEADERS)) + 1):
        wms_key = _cell_text(ws, row_index, COL_RETURN_WMS_ORDER_NO)
        ship_key = _cell_text(ws, row_index, COL_RETURN_SHIPMENT_NO)

        if wms_key:
            if wms_key in wms_to_shipment:
                if wms_to_shipment[wms_key] != ship_key:
                    raise InputError(
                        f"E12：WMS退单号 [{wms_key}] 已出现在物流单号 "
                        f"[{wms_to_shipment[wms_key]}]，当前行 {row_index} "
                        f"又出现在物流单号 [{ship_key}]。"
                    )
            else:
                wms_to_shipment[wms_key] = ship_key

        result.append(
            RawReturnRow(
                excel_row_num=row_index,
                shipment_no=ws.cell(row_index, COL_RETURN_SHIPMENT_NO).value,
                wms_order_no=ws.cell(row_index, COL_RETURN_WMS_ORDER_NO).value,
                sku=ws.cell(row_index, COL_RETURN_SKU).value,
                line_no=ws.cell(row_index, COL_RETURN_LINE_NO).value,
                qty=ws.cell(row_index, COL_RETURN_QTY).value,
            )
        )

    return result


def read_qc_inventory(ws: Any) -> list[RawInventoryRow]:
    """读取质检库存表，效期列记录单元格存储类型。"""
    _validate_header(ws, INVENTORY_HEADERS)

    result: list[RawInventoryRow] = []
    for row_index in range(2, _last_used_row(ws, len(INVENTORY_HEADERS)) + 1):
        expiry_cell = ws.cell(row_index, COL_INV_EXPIRY)
        result.append(
            RawInventoryRow(
                excel_row_num=row_index,
                shipment_no=ws.cell(row_index, COL_INV_SHIPMENT_NO).value,
                sku=ws.cell(row_index, COL_INV_SKU).value,
                qc=ws.cell(row_index, COL_INV_QC).value,
                lot_no=ws.cell(row_index, COL_INV_LOT_NO).value,
                expiry=expiry_cell.value,
                expiry_cell_kind=get_expiry_cell_kind(expiry_cell.value),
                qty=ws.cell(row_index, COL_INV_QTY).value,
            )
        )

    return result


# -----------------------------------------------------------------------------
# 效期单元格类型（对应 VBA 的 VarType(cell.Value) 判断）
# -----------------------------------------------------------------------------


def get_expiry_cell_kind(value: Any) -> str:
    """openpyxl 将 Excel 日期序列号还原为 datetime，与 VBA VarType=vbDate 对应。"""
    if isinstance(value, (datetime, date)):
        return CELL_KIND_EXCEL_DATE
    if isinstance(value, str):
        return CELL_KIND_TEXT
    if value is None:
        return CELL_KIND_BLANK
    return CELL_KIND_OTHER


# -----------------------------------------------------------------------------
# 表头校验（E12-①）
# -----------------------------------------------------------------------------


def _validate_header(ws: Any, expected_headers: list[str]) -> None:
    if ws is None:
        raise InputError("数据加载失败：工作表对象为空。")
    for col_index, expected in enumerate(expected_headers, start=1):
        actual = _cell_text(ws, 1, col_index)
        if actual != expected:
            raise InputError(
                f"表头校验失败：工作表 [{ws.title}] 第 {col_index} "
                f"列应为 [{expected}]，实际为 [{actual}]。"
            )


# -----------------------------------------------------------------------------
# 通用工具
# -----------------------------------------------------------------------------


def _cell_text(ws: Any, row: int, col: int) -> str:
    value = ws.cell(row, col).value
    return "" if value is None else str(value).strip()


def _last_used_row(ws: Any, col_count: int) -> int:
    """业务列范围内最后一个含非空值的行号；无数据时返回 1（对应 VBA Find 语义）。"""
    last = 1
    for row in ws.iter_rows(min_row=2, max_col=col_count):
        if any(cell.value is not None and str(cell.value) != "" for cell in row):
            last = row[0].row
    return last
