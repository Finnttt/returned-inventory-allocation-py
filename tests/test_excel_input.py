"""M03 数据加载单元测试（表头校验 E12-①、WMS退单号唯一性 E12-②、效期单元格类型）。"""

from datetime import datetime

import pytest
from openpyxl import Workbook

from returned_inventory.excel_input import (
    InputError,
    get_expiry_cell_kind,
    read_qc_inventory,
    read_return_orders,
)
from returned_inventory.models import (
    CELL_KIND_BLANK,
    CELL_KIND_EXCEL_DATE,
    CELL_KIND_OTHER,
    CELL_KIND_TEXT,
)


def make_return_ws(rows, headers=("物流单号", "WMS退单号", "SKU", "行号", "数量")):
    wb = Workbook()
    ws = wb.active
    ws.title = "输入_退单表"
    ws.append(list(headers))
    for row in rows:
        ws.append(list(row))
    return ws


def make_inventory_ws(rows, headers=("物流单号", "SKU", "QC情况", "批号", "效期", "数量")):
    wb = Workbook()
    ws = wb.active
    ws.title = "输入_质检库存表"
    ws.append(list(headers))
    for row in rows:
        ws.append(list(row))
    return ws


class TestReadReturnOrders:
    def test_reads_raw_values_without_transform(self):
        ws = make_return_ws([("SF001", "WMS001", "SKU1", "00001", 2)])
        rows = read_return_orders(ws)
        assert len(rows) == 1
        assert rows[0].excel_row_num == 2
        assert rows[0].line_no == "00001"

    def test_numeric_line_no_kept_as_is(self):
        ws = make_return_ws([("SF001", "WMS001", "SKU1", 1, 2)])
        rows = read_return_orders(ws)
        assert rows[0].line_no == 1  # 数值原样保留，由标准化层判非法

    def test_header_mismatch_raises(self):
        ws = make_return_ws([], headers=("物流单号", "WMS退单号", "sku", "行号", "数量"))
        with pytest.raises(InputError, match="表头校验失败"):
            read_return_orders(ws)

    def test_wms_cross_shipment_raises_e12(self):
        ws = make_return_ws(
            [
                ("SF001", "WMS001", "SKU1", "00001", 1),
                ("SF002", "WMS001", "SKU1", "00001", 1),
            ]
        )
        with pytest.raises(InputError, match="E12"):
            read_return_orders(ws)

    def test_same_wms_same_shipment_ok(self):
        ws = make_return_ws(
            [
                ("SF001", "WMS001", "SKU1", "00001", 1),
                ("SF001", "WMS001", "SKU2", "00002", 1),
            ]
        )
        assert len(read_return_orders(ws)) == 2


class TestReadQCInventory:
    def test_expiry_cell_kinds(self):
        ws = make_inventory_ws(
            [
                ("SF001", "SKU1", "ZP", "L01", datetime(2029, 1, 1), 5),
                ("SF001", "SKU1", "ZP", "L02", "2029/01/01", 5),
                ("SF001", "SKU1", "ZP", "L03", None, 5),
                ("SF001", "SKU1", "ZP", "L04", 45000, 5),
            ]
        )
        rows = read_qc_inventory(ws)
        kinds = [r.expiry_cell_kind for r in rows]
        assert kinds == [CELL_KIND_EXCEL_DATE, CELL_KIND_TEXT, CELL_KIND_BLANK, CELL_KIND_OTHER]

    def test_header_missing_column_raises(self):
        ws = make_inventory_ws([], headers=("物流单号", "SKU", "QC情况", "效期", "批号", "数量"))
        with pytest.raises(InputError, match="批号"):
            read_qc_inventory(ws)


class TestGetExpiryCellKind:
    @pytest.mark.parametrize(
        "value,expected",
        [
            (datetime(2029, 1, 1), CELL_KIND_EXCEL_DATE),
            ("2029/01/01", CELL_KIND_TEXT),
            (None, CELL_KIND_BLANK),
            (123, CELL_KIND_OTHER),
        ],
    )
    def test_kinds(self, value, expected):
        assert get_expiry_cell_kind(value) == expected
