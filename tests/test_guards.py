"""M10 工程守卫单元测试（对应 VBA modGuards.bas RunGuardsSelfTest 与需求 §6.5）。

覆盖：守恒断言成立/被破坏、撤销一致性（栈空精确相等 / 栈非空差值校验）、
E99Error 消息格式与字段、守卫在真实分配流程中的挂载（正常不触发 / 撤销被破坏时抛 E99）。
"""

import pytest

from returned_inventory.backtracking import allocate_shipment
from returned_inventory.guards import (
    CONTEXT_UNDO_CONSISTENCY,
    E99Error,
    assert_conservation,
    assert_undo_consistency,
    raise_e99,
)
from returned_inventory.ledger import InventoryLedger, build_ledger, new_undo_log
from returned_inventory.models import (
    Config,
    InventoryKey,
    NormalizedInventoryLine,
    NormalizedReturnLine,
)
from returned_inventory.sort_filter import build_static_plan, run_precheck
from returned_inventory.strategies import AllocationAttempt, AllocationAttemptDetail

SHIP = "SF001"
SKU = "SKU-A"
EXPIRY = "2029/01/01"
KEY = InventoryKey(SHIP, SKU, "ZP", "LA01", EXPIRY)


def _inv_line(qty=10, qc="ZP", lot_no="LA01", ship=SHIP, sku=SKU):
    return NormalizedInventoryLine(
        excel_row_num=2,
        shipment_no=ship,
        sku=sku,
        qc=qc,
        lot_no=lot_no,
        expiry=EXPIRY,
        qty=qty,
        qc_valid=True,
        expiry_valid=True,
        qty_valid=True,
        empty_fields="",
    )


class _Detail:
    """带 alloc_qty 属性的最小明细（assert_conservation 只依赖该字段）。"""

    def __init__(self, alloc_qty):
        self.alloc_qty = alloc_qty


def _attempt(qc="ZP", lot_no="LA01", alloc_qty=5):
    return AllocationAttempt(
        success=True,
        strategy_used="策略二",
        used_qc=qc,
        undo_log=new_undo_log(),
        details=[AllocationAttemptDetail(qc, lot_no, EXPIRY, alloc_qty)],
    )


class TestE99Error:
    """RaiseE99 消息格式（对应 VBA RunGuardsSelfTest 用例2）。"""

    def test_message_format_and_fields(self):
        with pytest.raises(E99Error) as exc_info:
            raise_e99("SF_GD_SELF02", "H_GD_SELF02", 100, 90, "自检")
        err = exc_info.value
        assert err.ship_no == "SF_GD_SELF02"
        assert err.sku == "H_GD_SELF02"
        assert err.expected == 100
        assert err.actual == 90
        assert err.context == "自检"
        msg = str(err)
        assert msg.startswith("[E99] 库存守恒异常：")
        for fragment in ("SF_GD_SELF02", "H_GD_SELF02", "100", "90", "自检"):
            assert fragment in msg


class TestAssertConservation:
    """库存守恒断言（对应 VBA RunGuardsSelfTest 用例1 与 §6.5.1）。"""

    def test_conservation_holds(self):
        """扣减 5、明细记 5：10 = 5 + 5，守恒成立。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        undo_log = new_undo_log()
        assert ledger.deduct(KEY, 5, undo_log)
        assert assert_conservation(snapshot, ledger, [_Detail(5)]) is True

    def test_deduct_without_detail_raises(self):
        """扣减了库存但明细漏记：守恒被破坏 → E99。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        ledger.deduct(KEY, 5, new_undo_log())
        with pytest.raises(E99Error) as exc_info:
            assert_conservation(snapshot, ledger, [])
        err = exc_info.value
        assert err.expected == 10 and err.actual == 5
        assert err.ship_no == SHIP and err.sku == SKU
        assert err.context == "AssertConservation"

    def test_detail_without_deduct_raises(self):
        """明细多记但未扣减库存：守恒被破坏 → E99。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        with pytest.raises(E99Error):
            assert_conservation(snapshot, ledger, [_Detail(3)])

    def test_none_snapshot_raises(self):
        """快照为 None 无法核验，按失败处理（对应 VBA 返回 False → 上层 RaiseE99）。"""
        ledger = build_ledger([_inv_line(qty=10)])
        with pytest.raises(E99Error):
            assert_conservation(None, ledger, [])

    def test_none_ledger_raises(self):
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        with pytest.raises(E99Error):
            assert_conservation(snapshot, None, [])


class TestAssertUndoConsistency:
    """撤销快照一致性校验（§6.5.2，对应 VBA AssertUndoConsistency）。"""

    def test_after_full_undo_passes(self):
        """扣减后完整撤销：账本与快照精确相等（栈空判定）。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        undo_log = new_undo_log()
        ledger.deduct(KEY, 5, undo_log)
        ledger.undo(undo_log)
        assert assert_undo_consistency(snapshot, None, ledger) is True

    def test_missing_undo_raises(self):
        """扣减后未撤销：账本与快照背离 → E99。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        ledger.deduct(KEY, 5, new_undo_log())
        with pytest.raises(E99Error) as exc_info:
            assert_undo_consistency(snapshot, None, ledger)
        assert exc_info.value.context == CONTEXT_UNDO_CONSISTENCY

    def test_partial_stack_passes(self):
        """栈内仍有未撤销条目：账本 = 快照 - 栈内分配明细之和。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        ledger.deduct(KEY, 5, new_undo_log())
        assert assert_undo_consistency(snapshot, [_attempt(alloc_qty=5)], ledger) is True

    def test_partial_stack_mismatch_raises(self):
        """栈内剩余条目与实际扣减不一致 → E99。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        ledger.deduct(KEY, 5, new_undo_log())
        with pytest.raises(E99Error):
            assert_undo_consistency(snapshot, [_attempt(alloc_qty=3)], ledger)

    def test_none_inputs_raise(self):
        ledger = build_ledger([_inv_line(qty=10)])
        snapshot = ledger.take_snapshot(SHIP, SKU)
        with pytest.raises(E99Error):
            assert_undo_consistency(None, None, ledger)
        with pytest.raises(E99Error):
            assert_undo_consistency(snapshot, None, None)


class TestGuardMountedInFlow:
    """守卫在真实分配流程中的挂载：正常流程不触发，撤销被破坏时立即抛 E99。"""

    @staticmethod
    def _build_tc21_h1():
        """TC-21 H000000001 场景（必然触发回溯）：退单 D=12+5×6，库存 ZP=20、QC=22。"""
        ship, sku = "SF3190000000016", "H000000001"
        rows = [
            NormalizedReturnLine(
                excel_row_num=i + 2,
                shipment_no=ship,
                wms_order_no="TK10000161",
                sku=sku,
                line_no=f"0000{i + 1}",
                qty=12 if i == 0 else 5,
                line_no_valid=True,
                qty_valid=True,
                empty_fields="",
            )
            for i in range(7)
        ]
        inv = [
            _inv_line(qty=8, qc="ZP", ship=ship, sku=sku),
            _inv_line(qty=12, qc="ZP", ship=ship, sku=sku),
            _inv_line(qty=12, qc="QC", ship=ship, sku=sku),
            _inv_line(qty=5, qc="QC", ship=ship, sku=sku),
            _inv_line(qty=5, qc="QC", ship=ship, sku=sku),
        ]
        ledger = build_ledger(inv)
        plan = build_static_plan(rows, ledger)
        precheck = run_precheck(plan, ledger)
        return ship, sku, ledger, plan, precheck

    def test_normal_flow_does_not_raise(self):
        """正常回溯流程（TC-21）：两个守卫挂载点均不触发。"""
        ship, sku, ledger, plan, precheck = self._build_tc21_h1()
        result = allocate_shipment(ship, [sku], {sku: plan}, {sku: precheck}, ledger, Config())
        assert result.group_results[0].success is True

    def test_broken_undo_triggers_e99(self, monkeypatch):
        """人为破坏 undo（只清日志不恢复库存）→ 回溯撤销后守卫立即抛 E99。"""
        ship, sku, ledger, plan, precheck = self._build_tc21_h1()

        def broken_undo(self, undo_log):
            undo_log.clear()  # 模拟撤销日志丢失：不恢复库存

        monkeypatch.setattr(InventoryLedger, "undo", broken_undo)
        with pytest.raises(E99Error) as exc_info:
            allocate_shipment(ship, [sku], {sku: plan}, {sku: precheck}, ledger, Config())
        err = exc_info.value
        assert err.context == CONTEXT_UNDO_CONSISTENCY
        assert err.ship_no == ship and err.sku == sku
