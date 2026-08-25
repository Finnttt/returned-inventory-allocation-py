"""M12 分配后校验单元测试（对应 VBA modPostValidate.bas 与需求 §4.3，TC-50）。

覆盖：正常分配结果通过、数量不一致、同一行两种 QC、整单回滚物流单号
不参与后校验、退单号状态与行状态聚合不一致、明细数据完整性各分支，
以及 assert_post_valid 的 E99 抛出行为。
"""

from dataclasses import dataclass

import pytest

from returned_inventory.guards import E99Error
from returned_inventory.models import (
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STATUS_UNALLOCATED,
    NormalizedReturnLine,
    WMSStatusEntry,
)
from returned_inventory.post_validate import (
    POST_ERR_DATA_MISMATCH,
    POST_ERR_QC_MISMATCH,
    POST_ERR_QTY_MISMATCH,
    POST_ERR_STATUS_MISMATCH,
    assert_post_valid,
    validate_post,
)

SHIP = "SF001"
WMS = "WMS001"
SKU = "SKU-A"
EXPIRY = "2029/01/01"


def _order(qty=5, ship=SHIP, wms=WMS, sku=SKU, line_no="00001"):
    return NormalizedReturnLine(
        excel_row_num=2,
        shipment_no=ship,
        wms_order_no=wms,
        sku=sku,
        line_no=line_no,
        qty=qty,
        line_no_valid=True,
        qty_valid=True,
        empty_fields="",
    )


@dataclass
class _FinalDetail:
    """M11 FinalResult 的 Detail_i_* 最小等价物：AllocationDetail 字段 + wms_order_status。"""

    shipment_no: str = SHIP
    wms_order_no: str = WMS
    sku: str = SKU
    line_no: str = "00001"
    order_qty: int = 5
    qc: str = "ZP"
    lot_no: str = "LA01"
    expiry: str = EXPIRY
    alloc_qty: int = 5
    line_status: str = STATUS_BATCH_IMPORT
    wms_order_status: str = STATUS_BATCH_IMPORT


def _summary(status=STATUS_BATCH_IMPORT, ship=SHIP, wms=WMS):
    return WMSStatusEntry(shipment_no=ship, wms_order_no=wms, status=status, reason="")


def _codes(result):
    return [issue.code for issue in result.issues]


class TestValidatePost:
    def test_normal_allocation_passes(self):
        result = validate_post([_order()], [_FinalDetail()], [_summary()])
        assert result.has_failures is False
        assert result.issue_count == 0

    def test_split_lot_same_qc_passes(self):
        # 同一行拆批号但 QC 一致：允许
        details = [
            _FinalDetail(alloc_qty=2, lot_no="LA01"),
            _FinalDetail(alloc_qty=3, lot_no="LB01"),
        ]
        result = validate_post([_order()], details, [_summary()])
        assert result.has_failures is False

    def test_qty_mismatch_fails(self):
        detail = _FinalDetail(alloc_qty=3)
        result = validate_post([_order(qty=5)], [detail], [_summary()])
        assert result.has_failures is True
        assert _codes(result) == [POST_ERR_QTY_MISMATCH]
        issue = result.issues[0]
        assert issue.shipment_no == SHIP
        assert issue.sku == SKU
        assert "分配量合计 3，退单量 5" in issue.message

    def test_missing_detail_fails_qty_mismatch(self):
        result = validate_post([_order()], [], [_summary()])
        assert _codes(result) == [POST_ERR_QTY_MISMATCH]

    def test_same_line_two_qc_fails(self):
        # 分配量合计正确（2+3=5），但同一行用了两种 QC
        details = [
            _FinalDetail(alloc_qty=2, qc="ZP"),
            _FinalDetail(alloc_qty=3, qc="NG"),
        ]
        result = validate_post([_order()], details, [_summary()])
        assert _codes(result) == [POST_ERR_QC_MISMATCH]
        assert "ZP,NG" in result.issues[0].message

    def test_rollback_shipment_skipped(self):
        # 整单回滚的物流单号：退单行无成功明细也不报数量不一致
        orders = [_order(), _order(ship="SF002", wms="WMS002")]
        details = [_FinalDetail()]
        summaries = [
            _summary(),
            _summary(status=STATUS_UNALLOCATED, ship="SF002", wms="WMS002"),
        ]
        result = validate_post(orders, details, summaries)
        assert result.has_failures is False

    def test_rollback_shipment_detail_is_data_mismatch(self):
        # 整单回滚的物流单号不应出现在成功分配明细中
        details = [_FinalDetail()]
        summaries = [_summary(status=STATUS_UNALLOCATED)]
        result = validate_post([], details, summaries)
        assert result.has_failures is True
        assert POST_ERR_DATA_MISMATCH in _codes(result)
        assert "整单回滚物流单号不应出现在成功分配明细中" in result.issues[0].message

    def test_status_aggregation_mismatch_fails(self):
        # 行状态聚合为「批量导入」，汇总表却写「手工操作」
        result = validate_post([_order()], [_FinalDetail()], [_summary(status=STATUS_MANUAL)])
        codes = _codes(result)
        assert POST_ERR_STATUS_MISMATCH in codes
        messages = [i.message for i in result.issues if i.code == POST_ERR_STATUS_MISMATCH]
        assert any("行状态聚合应为 批量导入，汇总表实际为 手工操作" in m for m in messages)

    def test_manual_line_aggregation_passes(self):
        # 任一行「手工操作」则聚合为「手工操作」
        detail = _FinalDetail(line_status=STATUS_MANUAL, wms_order_status=STATUS_MANUAL)
        result = validate_post([_order()], [detail], [_summary(status=STATUS_MANUAL)])
        assert result.has_failures is False

    def test_detail_line_status_vs_summary_mismatch(self):
        # 明细上的退单号状态与汇总表不一致
        detail = _FinalDetail(wms_order_status=STATUS_MANUAL)
        result = validate_post([_order()], [detail], [_summary()])
        codes = _codes(result)
        assert POST_ERR_STATUS_MISMATCH in codes
        messages = [i.message for i in result.issues if i.code == POST_ERR_STATUS_MISMATCH]
        assert any("明细退单号状态 手工操作，汇总退单号状态 批量导入" in m for m in messages)

    def test_illegal_line_status_fails(self):
        detail = _FinalDetail(line_status="莫名其妙", wms_order_status=STATUS_BATCH_IMPORT)
        result = validate_post([_order()], [detail], [_summary()])
        messages = [i.message for i in result.issues if i.code == POST_ERR_STATUS_MISMATCH]
        assert any("行状态非法：莫名其妙" in m for m in messages)

    def test_summary_missing_wms_status_fails(self):
        result = validate_post([_order()], [_FinalDetail()], [])
        messages = [i.message for i in result.issues if i.code == POST_ERR_STATUS_MISMATCH]
        assert any("汇总表缺少该 WMS 退单号状态" in m for m in messages)

    def test_detail_without_order_line_fails(self):
        detail = _FinalDetail(line_no="00099")
        result = validate_post([_order()], [detail], [_summary()])
        assert POST_ERR_DATA_MISMATCH in _codes(result)
        messages = [i.message for i in result.issues if i.code == POST_ERR_DATA_MISMATCH]
        assert any("成功分配明细找不到对应退单行" in m for m in messages)

    def test_detail_blank_key_field_fails(self):
        detail = _FinalDetail(sku=" ")
        result = validate_post([_order()], [detail], [_summary()])
        messages = [i.message for i in result.issues if i.code == POST_ERR_DATA_MISMATCH]
        assert any("成功分配明细关键字段为空" in m for m in messages)

    def test_detail_order_qty_mismatch_fails(self):
        detail = _FinalDetail(order_qty=4)
        result = validate_post([_order(qty=5)], [detail], [_summary()])
        messages = [i.message for i in result.issues if i.code == POST_ERR_DATA_MISMATCH]
        assert any("成功明细退单数量 4，输入退单数量 5" in m for m in messages)

    def test_detail_non_positive_alloc_qty_fails(self):
        # alloc_qty=0 不进入 QC 记录，但完整性检查要求分配数量必须大于 0
        detail = _FinalDetail(alloc_qty=0)
        result = validate_post([_order()], [detail], [_summary()])
        messages = [i.message for i in result.issues if i.code == POST_ERR_DATA_MISMATCH]
        assert any("成功分配明细分配数量必须大于 0" in m for m in messages)

    def test_detail_blank_lot_fails(self):
        detail = _FinalDetail(lot_no="")
        result = validate_post([_order()], [detail], [_summary()])
        messages = [i.message for i in result.issues if i.code == POST_ERR_DATA_MISMATCH]
        assert any("成功分配明细 QC/批号/效期不能为空" in m for m in messages)

    def test_multi_shipment_mixed_result(self):
        # 一个正常物流单号 + 一个数量不一致的物流单号：只报后者
        orders = [_order(), _order(qty=2, ship="SF002", wms="WMS002", line_no="00001")]
        details = [
            _FinalDetail(),
            _FinalDetail(shipment_no="SF002", wms_order_no="WMS002", order_qty=2, alloc_qty=1),
        ]
        summaries = [_summary(), _summary(ship="SF002", wms="WMS002")]
        result = validate_post(orders, details, summaries)
        assert _codes(result) == [POST_ERR_QTY_MISMATCH]
        assert result.issues[0].shipment_no == "SF002"


class TestAssertPostValid:
    def test_pass_on_success(self):
        result = validate_post([_order()], [_FinalDetail()], [_summary()])
        assert_post_valid(result)  # 不抛异常

    def test_raise_e99_on_failure(self):
        result = validate_post([_order(qty=5)], [_FinalDetail(alloc_qty=3)], [_summary()])
        with pytest.raises(E99Error) as exc_info:
            assert_post_valid(result)
        err = exc_info.value
        assert err.ship_no == SHIP
        assert err.sku == SKU
        assert POST_ERR_QTY_MISMATCH in err.context
        assert "分配量合计 3，退单量 5" in err.context
