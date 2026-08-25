"""M06 库存账本单元测试（对应 TC-56 与模块方案独立测试点）。"""

from returned_inventory.ledger import build_ledger, new_undo_log
from returned_inventory.models import InventoryKey, NormalizedInventoryLine


def _inv_line(
    shipment_no="SF001",
    sku="SKU-A",
    qc="ZP",
    lot_no="L01",
    expiry="2027/01/01",
    qty=10,
    qc_valid=True,
    expiry_valid=True,
    qty_valid=True,
    excel_row_num=2,
):
    return NormalizedInventoryLine(
        excel_row_num=excel_row_num,
        shipment_no=shipment_no,
        sku=sku,
        qc=qc,
        lot_no=lot_no,
        expiry=expiry,
        qty=qty,
        qc_valid=qc_valid,
        expiry_valid=expiry_valid,
        qty_valid=qty_valid,
        empty_fields="",
    )


KEY = InventoryKey("SF001", "SKU-A", "ZP", "L01", "2027/01/01")


class TestBuildLedger:
    """建账本：五元组汇总 + 非法行过滤。"""

    def test_same_five_tuple_aggregates(self):
        """相同五元组多行正确汇总（原始量与当前量都累加）。"""
        ledger = build_ledger([_inv_line(qty=10), _inv_line(qty=5), _inv_line(qty=3)])
        assert ledger.get_original_qty(KEY) == 18
        assert ledger.get_current_qty(KEY) == 18
        assert ledger.get_ledger_keys() == [KEY]

    def test_different_tuples_stay_separate(self):
        ledger = build_ledger(
            [_inv_line(qty=10), _inv_line(lot_no="L02", qty=7), _inv_line(qc="QC", qty=4)]
        )
        assert len(ledger.get_ledger_keys()) == 3
        assert ledger.query_qc_total("SF001", "SKU-A", "ZP") == 17

    def test_invalid_rows_skipped(self):
        """qc_valid/expiry_valid/qty_valid 任一为 False 的行不计入账本。"""
        ledger = build_ledger(
            [
                _inv_line(qty=10),
                _inv_line(qty=100, qc_valid=False),
                _inv_line(qty=100, expiry_valid=False),
                _inv_line(qty=100, qty_valid=False),
            ]
        )
        assert ledger.get_current_qty(KEY) == 10

    def test_empty_input(self):
        ledger = build_ledger([])
        assert ledger.get_ledger_keys() == []


class TestQueryAndRows:
    """query_qc_total 与 get_five_tuple_rows。"""

    def test_query_qc_total_sums_across_lots(self):
        ledger = build_ledger(
            [
                _inv_line(lot_no="L01", qty=10),
                _inv_line(lot_no="L02", expiry="2027/06/01", qty=5),
                _inv_line(qc="QC", qty=99),
                _inv_line(sku="SKU-B", qty=77),
            ]
        )
        assert ledger.query_qc_total("SF001", "SKU-A", "ZP") == 15
        assert ledger.query_qc_total("SF001", "SKU-A", "QC") == 99
        assert ledger.query_qc_total("SF001", "SKU-A", "NG") == 0

    def test_get_five_tuple_rows(self):
        ledger = build_ledger(
            [_inv_line(lot_no="L01", qty=10), _inv_line(lot_no="L02", qty=5)]
        )
        rows = ledger.get_five_tuple_rows("SF001", "SKU-A", "ZP")
        assert len(rows) == 2
        assert [r.lot_no for r in rows] == ["L01", "L02"]
        row = rows[0]
        assert row.key == KEY
        assert (row.original_qty, row.current_qty) == (10, 10)

    def test_get_five_tuple_rows_empty(self):
        ledger = build_ledger([_inv_line()])
        assert ledger.get_five_tuple_rows("SF001", "SKU-A", "NG") == []


class TestDeduct:
    """扣减：成功、库存不足、key 不存在。"""

    def test_deduct_success(self):
        ledger = build_ledger([_inv_line(qty=10)])
        undo_log = new_undo_log()
        assert ledger.deduct(KEY, 4, undo_log) is True
        assert ledger.get_current_qty(KEY) == 6
        assert ledger.get_original_qty(KEY) == 10
        assert undo_log == [(KEY, 4)]

    def test_deduct_insufficient_fails_and_ledger_unchanged(self):
        """扣减超过剩余量时失败且账本与 undo_log 均不变。"""
        ledger = build_ledger([_inv_line(qty=10)])
        undo_log = new_undo_log()
        assert ledger.deduct(KEY, 11, undo_log) is False
        assert ledger.get_current_qty(KEY) == 10
        assert ledger.get_original_qty(KEY) == 10
        assert undo_log == []

    def test_deduct_exact_remaining_succeeds(self):
        ledger = build_ledger([_inv_line(qty=10)])
        undo_log = new_undo_log()
        assert ledger.deduct(KEY, 10, undo_log) is True
        assert ledger.get_current_qty(KEY) == 0

    def test_deduct_missing_key_fails(self):
        ledger = build_ledger([_inv_line(qty=10)])
        undo_log = new_undo_log()
        missing = InventoryKey("SF001", "SKU-A", "ZP", "L99", "2027/01/01")
        assert ledger.deduct(missing, 1, undo_log) is False
        assert ledger.get_current_qty(KEY) == 10
        assert undo_log == []

    def test_get_qty_of_missing_key_is_zero(self):
        ledger = build_ledger([_inv_line(qty=10)])
        missing = InventoryKey("SF999", "SKU-Z", "QC", "L99", "2030/01/01")
        assert ledger.get_current_qty(missing) == 0
        assert ledger.get_original_qty(missing) == 0


class TestUndo:
    """undo：LIFO 还原，完成后账本与扣减前完全一致，日志清空。"""

    def test_undo_restores_exactly(self):
        ledger = build_ledger([_inv_line(qty=10), _inv_line(lot_no="L02", qty=5)])
        key2 = InventoryKey("SF001", "SKU-A", "ZP", "L02", "2027/01/01")
        before = [(k, ledger.get_original_qty(k), ledger.get_current_qty(k)) for k in ledger.get_ledger_keys()]

        undo_log = new_undo_log()
        ledger.deduct(KEY, 3, undo_log)
        ledger.deduct(key2, 5, undo_log)
        ledger.deduct(KEY, 2, undo_log)
        ledger.undo(undo_log)

        after = [(k, ledger.get_original_qty(k), ledger.get_current_qty(k)) for k in ledger.get_ledger_keys()]
        assert after == before
        assert undo_log == []

    def test_undo_partial_failure_leaves_success_revertible(self):
        """失败的扣减不入日志，undo 只还原成功记录。"""
        ledger = build_ledger([_inv_line(qty=10)])
        undo_log = new_undo_log()
        ledger.deduct(KEY, 4, undo_log)
        ledger.deduct(KEY, 99, undo_log)  # 失败，不入日志
        ledger.undo(undo_log)
        assert ledger.get_current_qty(KEY) == 10
        assert undo_log == []

    def test_undo_on_empty_log_is_noop(self):
        ledger = build_ledger([_inv_line(qty=10)])
        ledger.undo(new_undo_log())
        assert ledger.get_current_qty(KEY) == 10


class TestSnapshot:
    """快照：范围正确、深拷贝独立、与 query_qc_total 一致。"""

    def test_snapshot_scope_and_values(self):
        ledger = build_ledger(
            [
                _inv_line(qty=10),
                _inv_line(qc="QC", qty=4),
                _inv_line(sku="SKU-B", qty=77),
            ]
        )
        snap = ledger.take_snapshot("SF001", "SKU-A")
        assert snap.get_snapshot_current_qty(KEY) == 10
        assert snap.get_snapshot_current_qty(InventoryKey("SF001", "SKU-A", "QC", "L01", "2027/01/01")) == 4
        # SKU-B 不在快照范围内
        assert len(snap.get_ledger_keys()) == 2

    def test_snapshot_is_deep_copy(self):
        """账本后续 deduct/undo 不影响快照。"""
        ledger = build_ledger([_inv_line(qty=10)])
        snap = ledger.take_snapshot("SF001", "SKU-A")
        undo_log = new_undo_log()
        ledger.deduct(KEY, 6, undo_log)
        assert snap.get_snapshot_current_qty(KEY) == 10
        ledger.undo(undo_log)
        assert snap.get_snapshot_current_qty(KEY) == 10

    def test_snapshot_consistent_with_query_qc_total(self):
        ledger = build_ledger(
            [_inv_line(lot_no="L01", qty=10), _inv_line(lot_no="L02", qty=5)]
        )
        snap = ledger.take_snapshot("SF001", "SKU-A")
        snap_total = sum(snap.get_snapshot_current_qty(k) for k in snap.get_ledger_keys())
        assert snap_total == ledger.query_qc_total("SF001", "SKU-A", "ZP")

    def test_snapshot_empty_scope(self):
        ledger = build_ledger([_inv_line()])
        snap = ledger.take_snapshot("SF999", "SKU-Z")
        assert snap.get_ledger_keys() == []
