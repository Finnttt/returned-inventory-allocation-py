"""M08 分配策略单元测试（对应 TC-01/02/03/09/10 与需求 §4.2.4 / §4.2.5 / §6.3.2）。"""

from returned_inventory.ledger import InventoryLedger
from returned_inventory.models import (
    STRATEGY_FAILED,
    STRATEGY_ONE,
    STRATEGY_THREE,
    STRATEGY_TWO,
    InventoryRow,
)
from returned_inventory.strategies import (
    AllocationAttemptDetail,
    compare_by_priority,
    try_allocate,
)

SHIP = "SF001"
SKU = "SKU-A"


def _row(qc="ZP", lot_no="L01", expiry="2027/01/01", qty=10):
    """构造候选行（CandidateRow = InventoryRow），original/current 数量相同。"""
    return InventoryRow(
        shipment_no=SHIP,
        sku=SKU,
        qc=qc,
        lot_no=lot_no,
        expiry=expiry,
        original_qty=qty,
        current_qty=qty,
    )


def _ledger_for(rows):
    """按候选行当前数量构造账本（绕过 M04/M06 输入链路，直接测策略）。"""
    return InventoryLedger({r.key: [r.original_qty, r.current_qty] for r in rows})


def _qty(ledger, row):
    return ledger.get_current_qty(row.key)


class TestCompareByPriority:
    """TC-10：三级比较规则（QC优先级 → 效期降序 → 批号升序）。"""

    def test_qc_priority_zp_over_qc_over_ng(self):
        zp = _row(qc="ZP")
        qc = _row(qc="QC")
        ng = _row(qc="NG")
        assert compare_by_priority(zp, qc) < 0
        assert compare_by_priority(qc, ng) < 0
        assert compare_by_priority(ng, zp) > 0

    def test_unknown_qc_ranks_last(self):
        assert compare_by_priority(_row(qc="NG"), _row(qc="XX")) < 0

    def test_same_qc_later_expiry_first(self):
        early = _row(expiry="2027/01/01")
        late = _row(expiry="2029/12/31")
        assert compare_by_priority(late, early) < 0
        assert compare_by_priority(early, late) > 0

    def test_same_qc_same_expiry_smaller_lot_first(self):
        small = _row(lot_no="A01")
        big = _row(lot_no="B02")
        assert compare_by_priority(small, big) < 0
        assert compare_by_priority(big, small) > 0

    def test_identical_key_fields_equal(self):
        assert compare_by_priority(_row(), _row()) == 0

    def test_qc_beats_expiry_and_lot(self):
        """QC 优先级高于效期与批号：NG 效期再晚也排在 ZP 之后。"""
        zp_early = _row(qc="ZP", expiry="2027/01/01", lot_no="Z99")
        ng_late = _row(qc="NG", expiry="2030/01/01", lot_no="A01")
        assert compare_by_priority(zp_early, ng_late) < 0


class TestBoundary:
    """边界：需求无效或候选池为空，直接失败且不动账本。"""

    def test_empty_pool(self):
        attempt = try_allocate([], 5, _ledger_for([]))
        assert attempt.success is False
        assert attempt.strategy_used == STRATEGY_FAILED
        assert attempt.detail_count == 0
        assert attempt.undo_log == []

    def test_non_positive_demand(self):
        row = _row(qty=10)
        ledger = _ledger_for([row])
        attempt = try_allocate([row], 0, ledger)
        assert attempt.success is False
        assert _qty(ledger, row) == 10


class TestStrategyOne:
    """TC-01：策略一精确匹配。"""

    def test_exact_match(self):
        row = _row(qty=5)
        ledger = _ledger_for([row])
        attempt = try_allocate([row], 5, ledger)
        assert attempt.success is True
        assert attempt.strategy_used == STRATEGY_ONE
        assert attempt.used_qc == "ZP"
        assert attempt.details == [AllocationAttemptDetail("ZP", "L01", "2027/01/01", 5)]
        assert _qty(ledger, row) == 0

    def test_tie_prefers_zp_over_qc_over_ng(self):
        zp = _row(qc="ZP", lot_no="B01", qty=5)
        qc = _row(qc="QC", lot_no="A01", qty=5)
        ng = _row(qc="NG", lot_no="A01", expiry="2030/01/01", qty=5)
        ledger = _ledger_for([zp, qc, ng])
        attempt = try_allocate([ng, qc, zp], 5, ledger)
        assert attempt.strategy_used == STRATEGY_ONE
        assert attempt.details[0].qc == "ZP"
        assert _qty(ledger, zp) == 0
        assert _qty(ledger, qc) == 5
        assert _qty(ledger, ng) == 5

    def test_tie_same_qc_prefers_later_expiry(self):
        early = _row(expiry="2027/01/01", qty=5)
        late = _row(lot_no="L02", expiry="2029/01/01", qty=5)
        ledger = _ledger_for([early, late])
        attempt = try_allocate([early, late], 5, ledger)
        assert attempt.details[0].expiry == "2029/01/01"
        assert _qty(ledger, late) == 0
        assert _qty(ledger, early) == 5

    def test_tie_same_qc_same_expiry_prefers_smaller_lot(self):
        big = _row(lot_no="B02", qty=5)
        small = _row(lot_no="A01", qty=5)
        ledger = _ledger_for([big, small])
        attempt = try_allocate([big, small], 5, ledger)
        assert attempt.details[0].lot_no == "A01"
        assert _qty(ledger, small) == 0
        assert _qty(ledger, big) == 5


class TestStrategyTwo:
    """TC-02：策略二最接近（最小 sufficient）匹配。"""

    def test_picks_smallest_sufficient(self):
        small = _row(lot_no="L01", qty=8)
        large = _row(lot_no="L02", qty=100)
        ledger = _ledger_for([small, large])
        attempt = try_allocate([small, large], 5, ledger)
        assert attempt.success is True
        assert attempt.strategy_used == STRATEGY_TWO
        assert attempt.details == [AllocationAttemptDetail("ZP", "L01", "2027/01/01", 5)]

    def test_remainder_stays_in_ledger(self):
        """扣减量 = 需求量，该行剩余库存保留（供后续行使用）。"""
        row = _row(qty=8)
        ledger = _ledger_for([row])
        attempt = try_allocate([row], 5, ledger)
        assert attempt.success is True
        assert _qty(ledger, row) == 3

    def test_exact_match_wins_over_sufficient(self):
        """qty == D 的行由策略一命中，不落入策略二。"""
        exact = _row(lot_no="L01", qty=5)
        sufficient = _row(lot_no="L02", qty=6)
        ledger = _ledger_for([exact, sufficient])
        attempt = try_allocate([sufficient, exact], 5, ledger)
        assert attempt.strategy_used == STRATEGY_ONE
        assert attempt.details[0].lot_no == "L01"

    def test_tie_same_qty_prefers_priority_order(self):
        """同样接近（qty 相同）时按 ZP>QC>NG → 效期晚 → 批号小。"""
        qc_row = _row(qc="QC", lot_no="A01", qty=8)
        zp_row = _row(qc="ZP", lot_no="B01", qty=8)
        ledger = _ledger_for([qc_row, zp_row])
        attempt = try_allocate([qc_row, zp_row], 5, ledger)
        assert attempt.strategy_used == STRATEGY_TWO
        assert attempt.details[0].qc == "ZP"
        assert _qty(ledger, zp_row) == 3
        assert _qty(ledger, qc_row) == 8


class TestStrategyThree:
    """TC-03/TC-09：策略三多行拼凑。"""

    def test_spec_example_demand_12(self):
        """需求 §4.2.4 示例：需求 12，第一步锁定 ZP，第二步平局按效期更晚
        选 W-2029/02/01，结果 ZP:L-2029/01/01-9 + ZP:W-2029/02/01-3。"""
        rows = [
            _row(qc="ZP", lot_no="L", expiry="2029/01/01", qty=9),
            _row(qc="ZP", lot_no="Q", expiry="2029/01/01", qty=7),
            _row(qc="ZP", lot_no="W", expiry="2029/01/01", qty=3),
            _row(qc="ZP", lot_no="W", expiry="2029/02/01", qty=3),
            _row(qc="QC", lot_no="L", expiry="2029/01/01", qty=9),
            _row(qc="QC", lot_no="Q", expiry="2029/01/01", qty=4),
            _row(qc="QC", lot_no="W", expiry="2029/01/01", qty=2),
        ]
        ledger = _ledger_for(rows)
        attempt = try_allocate(rows, 12, ledger)

        assert attempt.success is True
        assert attempt.strategy_used == STRATEGY_THREE
        assert attempt.used_qc == "ZP"
        assert attempt.details == [
            AllocationAttemptDetail("ZP", "L", "2029/01/01", 9),
            AllocationAttemptDetail("ZP", "W", "2029/02/01", 3),
        ]
        # 两个被命中的五元组扣到 0，其余行不动
        assert _qty(ledger, rows[0]) == 0
        assert _qty(ledger, rows[3]) == 0
        assert _qty(ledger, rows[1]) == 7
        assert _qty(ledger, rows[2]) == 3
        assert _qty(ledger, rows[4]) == 9

    def test_does_not_cross_qc_and_rolls_back(self):
        """策略三不跨 QC：锁定 QC 总量不足时撤销全部扣减，账本恢复原状。

        需求 10：QC:8 距 10 最近（锁定 QC），但 QC 总量只有 8 < 10，
        拼凑到第二步无行可选 → 失败回滚；ZP 侧 6+3=9 同样不够，验证不跨 QC 补足。
        """
        zp1 = _row(qc="ZP", lot_no="L01", qty=6)
        zp2 = _row(qc="ZP", lot_no="L02", qty=3)
        qc1 = _row(qc="QC", lot_no="L01", qty=8)
        ledger = _ledger_for([zp1, zp2, qc1])
        before = {r.key: _qty(ledger, r) for r in [zp1, zp2, qc1]}

        attempt = try_allocate([zp1, zp2, qc1], 10, ledger)

        assert attempt.success is False
        assert attempt.strategy_used == STRATEGY_FAILED
        assert attempt.detail_count == 0
        for r in [zp1, zp2, qc1]:
            assert _qty(ledger, r) == before[r.key]

    def test_partial_deduction_on_last_step(self):
        """允许最后一步部分扣减：选中行 qty > 剩余需求时只扣剩余量。"""
        r1 = _row(lot_no="L01", qty=5)
        r2 = _row(lot_no="L02", qty=4)
        ledger = _ledger_for([r1, r2])
        attempt = try_allocate([r1, r2], 7, ledger)

        assert attempt.success is True
        assert attempt.strategy_used == STRATEGY_THREE
        assert attempt.details == [
            AllocationAttemptDetail("ZP", "L01", "2027/01/01", 5),
            AllocationAttemptDetail("ZP", "L02", "2027/01/01", 2),
        ]
        assert _qty(ledger, r1) == 0
        assert _qty(ledger, r2) == 2  # 部分扣减后剩余库存保留

    def test_next_step_tie_prefers_covering_row(self):
        """后续步骤平局：距离相同时优先选 qty >= 剩余需求的行（一步完成）。

        需求 10：第一步 ZP:6（唯一最大），剩余 4；
        ZP 内剩 qty=2 与 qty=4，距 4 的距离为 2 与 0 —— 不构成平局。
        改用 qty=2 与 qty=6：距 4 都是 2，优先选能一步完成的 qty=6（部分扣减）。
        """
        r1 = _row(lot_no="L01", qty=6)
        r2 = _row(lot_no="L02", qty=2)
        r3 = _row(lot_no="L03", expiry="2028/01/01", qty=6)
        ledger = _ledger_for([r1, r2, r3])
        # 需求 10：第一步 r1/r3 距 10 都是 4，平局按效期更晚选 r3；
        # 剩余 4：r1(6) 与 r2(2) 距 4 都是 2，优先选覆盖的 r1，部分扣 4。
        attempt = try_allocate([r1, r2, r3], 10, ledger)

        assert attempt.success is True
        assert attempt.details == [
            AllocationAttemptDetail("ZP", "L03", "2028/01/01", 6),
            AllocationAttemptDetail("ZP", "L01", "2027/01/01", 4),
        ]
        assert _qty(ledger, r3) == 0
        assert _qty(ledger, r1) == 2
        assert _qty(ledger, r2) == 2

    def test_total_insufficient_fails_with_ledger_unchanged(self):
        """所有 QC 总量都不足需求：全失败，账本不变。"""
        r1 = _row(qc="ZP", lot_no="L01", qty=3)
        r2 = _row(qc="QC", lot_no="L01", qty=4)
        ledger = _ledger_for([r1, r2])
        attempt = try_allocate([r1, r2], 10, ledger)

        assert attempt.success is False
        assert _qty(ledger, r1) == 3
        assert _qty(ledger, r2) == 4

    def test_zero_qty_rows_ignored(self):
        """current_qty = 0 的行（已被此前行扣空）不参与策略三选择。"""
        empty = _row(lot_no="L01", qty=0)
        r2 = _row(lot_no="L02", qty=4)
        r3 = _row(lot_no="L03", qty=3)
        ledger = _ledger_for([empty, r2, r3])
        attempt = try_allocate([empty, r2, r3], 7, ledger)

        assert attempt.success is True
        assert attempt.strategy_used == STRATEGY_THREE
        assert [d.lot_no for d in attempt.details] == ["L02", "L03"]
        assert _qty(ledger, empty) == 0
