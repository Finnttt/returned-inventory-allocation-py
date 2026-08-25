"""M09 回溯分配引擎单元测试（对应 TC-21/22/24 与需求 §4.2 / §4.2.7 / §6.3.1.3）。

覆盖：无需回溯直分成功、回溯触发并成功（TC-21）、回溯路径穷尽 E09（TC-22）、
回溯超限 E10 + 跨 SKU 短路连带回滚（TC-24）、预检测A/B 命中短路、行状态判定、
调试日志三档行为与 E09 失败子类型。
"""

from returned_inventory.backtracking import (
    FAIL_SUB_CASCADE_CROSS_SKU,
    FAIL_SUB_CASCADE_SAME_SKU,
    FAIL_SUB_NO_AVAILABLE_QC,
    FAIL_SUB_PATH_EXHAUSTED,
    FAIL_SUB_PRECHECK_A,
    FAIL_SUB_PRECHECK_B,
    PRECHECK_HIT_A,
    PRECHECK_HIT_B,
    PROCESS_LINE_STATUS_ATTEMPT_FAIL,
    PROCESS_LINE_STATUS_ATTEMPT_OK,
    PROCESS_LINE_STATUS_REVOKE,
    allocate_shipment,
)
from returned_inventory.ledger import build_ledger
from returned_inventory.models import (
    DEBUG_LEVEL_DETAIL,
    DEBUG_LEVEL_OFF,
    DEBUG_LEVEL_SIMPLE,
    ERR_E09,
    ERR_E10,
    ERROR_CASCADE_ROLLBACK,
    LINE_STATUS_FAILED,
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STRATEGY_ONE,
    STRATEGY_THREE,
    STRATEGY_TWO,
    Config,
    NormalizedInventoryLine,
    NormalizedReturnLine,
)
from returned_inventory.sort_filter import build_static_plan, run_precheck

EXPIRY = "2029/01/01"


def _ret(ship, wms, sku, line_no, qty, row=2):
    return NormalizedReturnLine(
        excel_row_num=row,
        shipment_no=ship,
        wms_order_no=wms,
        sku=sku,
        line_no=line_no,
        qty=qty,
        line_no_valid=True,
        qty_valid=True,
        empty_fields="",
    )


def _inv(ship, sku, qc, qty, lot_no="LA01", row=2):
    return NormalizedInventoryLine(
        excel_row_num=row,
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


def _cfg(max_backtrack=200, level=DEBUG_LEVEL_OFF):
    return Config(max_backtrack_count=max_backtrack, debug_log_level=level)


def _run(ship, sku_list, rows_by_sku, inv_lines, cfg=None):
    """构建账本 → 各 SKU 静态计划与预检测 → allocate_shipment，返回 (结果, 账本)。"""
    ledger = build_ledger(inv_lines)
    plan_map = {}
    precheck_map = {}
    for sku in sku_list:
        plan = build_static_plan(rows_by_sku[sku], ledger)
        plan_map[sku] = plan
        precheck_map[sku] = run_precheck(plan, ledger)
    result = allocate_shipment(ship, sku_list, plan_map, precheck_map, ledger, cfg or _cfg())
    return result, ledger


# -----------------------------------------------------------------------------
# TC-21 场景数据：H000000001，退单 D=12+5×6，库存 ZP=20（8+12）、QC=22（12+5+5）
# -----------------------------------------------------------------------------

TC21_SHIP = "SF3190000000016"
TC21_SKU = "H000000001"


def _tc21_rows():
    return [
        _ret(TC21_SHIP, "TK10000161", TC21_SKU, "00001", 12, row=2),
        *[
            _ret(TC21_SHIP, "TK10000161", TC21_SKU, f"0000{i}", 5, row=i + 1)
            for i in range(2, 8)
        ],
    ]


def _tc21_inv():
    return [
        _inv(TC21_SHIP, TC21_SKU, "ZP", 8, row=2),
        _inv(TC21_SHIP, TC21_SKU, "ZP", 12, row=3),
        _inv(TC21_SHIP, TC21_SKU, "QC", 12, row=4),
        _inv(TC21_SHIP, TC21_SKU, "QC", 5, row=5),
        _inv(TC21_SHIP, TC21_SKU, "QC", 5, row=6),
    ]


# -----------------------------------------------------------------------------
# TC-22/24 场景数据：退单 D=6×3 + 4×6，库存 ZP=23、QC=13、NG=6（结构不可行）
# -----------------------------------------------------------------------------

TC22_SHIP = "SF3190000000027"
TC22_SKU = "H000000001"


def _tc22_rows(ship=TC22_SHIP, sku=TC22_SKU):
    return [
        *[_ret(ship, "TK10000271", sku, f"0000{i}", 6, row=i + 1) for i in range(1, 4)],
        *[_ret(ship, "TK10000271", sku, f"0000{i}", 4, row=i + 1) for i in range(4, 10)],
    ]


def _tc22_inv(ship=TC22_SHIP, sku=TC22_SKU):
    return [
        _inv(ship, sku, "ZP", 23, row=2),
        _inv(ship, sku, "QC", 13, row=3),
        _inv(ship, sku, "NG", 6, row=4),
    ]


class TestDirectSuccess:
    """无需回溯直分成功：backtrack_count=0，行状态与策略记录正确。"""

    def test_single_row_strategy_one(self):
        """单行 D=1，NG=1 精确匹配 → 策略一，批量导入，账本耗尽。"""
        ship, sku = "SF001", "SKU-A"
        rows = [_ret(ship, "WMS1", sku, "00001", 1)]
        inv = [_inv(ship, sku, "NG", 1, lot_no="LB01")]
        result, ledger = _run(ship, [sku], {sku: rows}, inv)

        group = result.group_results[0]
        assert group.success is True
        assert group.error_code == ""
        assert group.stats.backtrack_count == 0
        assert group.stats.precheck_hit == ""
        assert len(group.details) == 1
        d = group.details[0]
        assert (d.qc, d.alloc_qty, d.line_status, d.strategy_used) == (
            "NG", 1, STATUS_BATCH_IMPORT, STRATEGY_ONE,
        )
        assert ledger.query_qc_total(ship, sku, "NG") == 0

    def test_two_rows_strategy_two_then_one(self):
        """D={5,3}，ZP=8：首行策略二部分扣减（剩余保留），末行策略一精确耗尽。"""
        ship, sku = "SF002", "SKU-B"
        rows = [
            _ret(ship, "WMS1", sku, "00001", 5),
            _ret(ship, "WMS1", sku, "00002", 3),
        ]
        inv = [_inv(ship, sku, "ZP", 8)]
        result, ledger = _run(ship, [sku], {sku: rows}, inv)

        group = result.group_results[0]
        assert group.success is True
        assert group.stats.backtrack_count == 0
        by_line = {d.line_no: d for d in group.details}
        assert by_line["00001"].strategy_used == STRATEGY_TWO
        assert by_line["00002"].strategy_used == STRATEGY_ONE
        assert all(d.line_status == STATUS_BATCH_IMPORT for d in group.details)
        assert ledger.query_qc_total(ship, sku, "ZP") == 0

    def test_stats_produced_with_debug_off(self):
        """调试日志关闭：不产生任何事件，但 GroupStats 统计照常（§6.5.4 强制下限）。"""
        ship, sku = "SF003", "SKU-C"
        rows = [_ret(ship, "WMS1", sku, "00001", 1)]
        inv = [_inv(ship, sku, "NG", 1, lot_no="LB01")]
        result, _ = _run(ship, [sku], {sku: rows}, inv, _cfg(level=DEBUG_LEVEL_OFF))
        group = result.group_results[0]
        assert group.events == []
        assert group.stats.backtrack_count == 0


class TestBacktrackSuccessTC21:
    """TC-21：当前路径失败 → 回溯改选 QC 后成功（回溯计数=4，单行撤销-重试口径）。

    初始路径 00001 选 ZP 导致 00005 候选池为空；逐级回退 00004→00003→00002→00001
    （中间三行唯一可用 QC 已尝试过，重试仍失败，自然继续回退），00001 排除 ZP 改选
    QC 后全组成功。00002 回溯后以清空的已尝试列表重选 QC——证明「target 之后所有行
    的已尝试QC记录已清除」（§4.2.7 步骤 4，漏掉会漏解）。
    """

    def test_backtrack_success_final_state(self):
        result, ledger = _run(
            TC21_SHIP, [TC21_SKU], {TC21_SKU: _tc21_rows()}, _tc21_inv()
        )
        group = result.group_results[0]
        assert group.success is True
        assert group.error_code == ""
        assert group.stats.backtrack_count == 4
        assert group.stats.precheck_hit == ""
        assert len(group.details) == 7

        by_line = {d.line_no: d for d in group.details}
        expected = {
            "00001": ("QC", STRATEGY_TWO),
            "00002": ("QC", STRATEGY_TWO),
            "00003": ("QC", STRATEGY_ONE),
            "00004": ("ZP", STRATEGY_TWO),
            "00005": ("ZP", STRATEGY_TWO),
            "00006": ("ZP", STRATEGY_TWO),
            "00007": ("ZP", STRATEGY_ONE),
        }
        for line_no, (qc, strategy) in expected.items():
            d = by_line[line_no]
            assert d.qc == qc, line_no
            assert d.strategy_used == strategy, line_no
            assert d.line_status == STATUS_BATCH_IMPORT, line_no

        # 终态：ZP/QC 库存全部耗尽（E08 守恒：退单 42 = 库存 42）
        assert ledger.query_qc_total(TC21_SHIP, TC21_SKU, "ZP") == 0
        assert ledger.query_qc_total(TC21_SHIP, TC21_SKU, "QC") == 0

    def test_simple_log_final_events(self):
        """简版：每行 1 条最终结果事件，回溯次数/重试标记/选用QC 正确。"""
        result, _ = _run(
            TC21_SHIP, [TC21_SKU], {TC21_SKU: _tc21_rows()}, _tc21_inv(),
            _cfg(level=DEBUG_LEVEL_SIMPLE),
        )
        events = result.group_results[0].events
        assert len(events) == 7
        assert all(e.is_final_result and not e.is_revoked for e in events)
        assert all(e.backtrack_no == 4 and e.is_backtrack_retry == "是" for e in events)

        first = events[0]  # 处理序 1 = 行 00001
        assert first.line_no == "00001"
        assert first.process_order == "1"
        assert first.strategy_used == STRATEGY_TWO
        assert first.used_qc == "QC"
        assert first.candidate_qc_count == "2"  # 初始可用QC数（静态值）
        assert first.dynamic_next_min_qty == "5"
        assert first.line_status == STATUS_BATCH_IMPORT
        assert first.error_code == "" and first.fail_sub_type == ""

    def test_detail_log_process_events(self):
        """详细：过程事件（尝试成功/失败/回溯撤销）在前，最终结果事件在后。"""
        result, _ = _run(
            TC21_SHIP, [TC21_SKU], {TC21_SKU: _tc21_rows()}, _tc21_inv(),
            _cfg(level=DEBUG_LEVEL_DETAIL),
        )
        events = result.group_results[0].events
        # 尝试事件 15（初始 5 = 4 成功 + 1 失败；回退途中重试失败 3；改选后成功 7）
        # + 撤销事件 4 = 19 过程事件 + 7 最终结果事件
        assert len(events) == 26
        assert all(not e.is_final_result for e in events[:19])
        assert all(e.is_final_result for e in events[19:])

        revokes = [e for e in events if e.line_status == PROCESS_LINE_STATUS_REVOKE]
        assert len(revokes) == 4
        assert all(e.is_revoked for e in revokes)
        # 逐级回退：依次撤销处理序 4/3/2/1（行 00004/00003/00002/00001）
        assert [e.process_order for e in revokes] == ["4", "3", "2", "1"]

        fails = [e for e in events if e.line_status == PROCESS_LINE_STATUS_ATTEMPT_FAIL]
        assert len(fails) == 4
        assert all(e.error_code == ERR_E09 for e in fails)
        assert all(e.fail_sub_type == FAIL_SUB_NO_AVAILABLE_QC for e in fails)

        # 行 00001 重试时 ZP 已被排除（已尝试QC记录生效），候选仅剩 QC
        retries = [
            e for e in events
            if e.line_status == PROCESS_LINE_STATUS_ATTEMPT_OK and e.line_no == "00001"
        ]
        assert [e.excluded_qc_list for e in retries] == ["", "ZP"]
        assert retries[1].candidate_qc_count == "1"


class TestE09PathExhaustedTC22:
    """TC-22：15 条路径全部失败 → 回溯路径穷尽 E09（实现计步 59），账本完整还原。"""

    def test_e09_result_and_ledger_restored(self):
        result, ledger = _run(
            TC22_SHIP, [TC22_SKU], {TC22_SKU: _tc22_rows()}, _tc22_inv()
        )
        group = result.group_results[0]
        assert group.success is False
        assert group.error_code == ERR_E09
        assert group.stats.backtrack_count == 59
        assert group.stats.precheck_hit == ""
        assert group.details == []
        # E09 穷尽时选择栈已逐级清空：账本与初始状态一致
        assert ledger.query_qc_total(TC22_SHIP, TC22_SKU, "ZP") == 23
        assert ledger.query_qc_total(TC22_SHIP, TC22_SKU, "QC") == 13
        assert ledger.query_qc_total(TC22_SHIP, TC22_SKU, "NG") == 6

    def test_e09_final_events_simple(self):
        """简版失败事件：首行子类型=回溯路径穷尽，其后各行=连带回滚—同SKU未到达行。

        分支行为逐行对齐当前 VBA BT_FillFailureDebugFields（与 TC-22 文档中较早的
        冻结日志表叙述不同，见 backtracking.py 模块头偏差说明第 4 条）。
        """
        result, _ = _run(
            TC22_SHIP, [TC22_SKU], {TC22_SKU: _tc22_rows()}, _tc22_inv(),
            _cfg(level=DEBUG_LEVEL_SIMPLE),
        )
        events = result.group_results[0].events
        assert len(events) == 9
        assert all(e.is_final_result for e in events)
        assert all(e.error_code == ERR_E09 for e in events)
        assert all(e.line_status == LINE_STATUS_FAILED for e in events)
        assert all(e.backtrack_no == 59 and e.is_backtrack_retry == "是" for e in events)
        assert events[0].fail_sub_type == FAIL_SUB_PATH_EXHAUSTED
        for e in events[1:]:
            assert e.fail_sub_type == FAIL_SUB_CASCADE_SAME_SKU
            assert e.process_order == "-"

    def test_e09_stats_without_debug_log(self):
        """关闭调试日志：事件为空，回溯统计仍为 59（§6.5.4 强制下限）。"""
        result, _ = _run(
            TC22_SHIP, [TC22_SKU], {TC22_SKU: _tc22_rows()}, _tc22_inv(),
            _cfg(level=DEBUG_LEVEL_OFF),
        )
        group = result.group_results[0]
        assert group.events == []
        assert group.stats.backtrack_count == 59


class TestE10OverLimitTC24:
    """TC-24：max=10 时第 11 次回溯超限 → E10；H2 从未执行，连带回滚（§4.2 短路）。"""

    SHIP = "SF3190000000028"
    SKU_H1 = "H000000001"
    SKU_H2 = "H000000002"

    def _data(self):
        rows_h1 = [
            _ret(self.SHIP, "TK10000281", self.SKU_H1, "00001", 6, row=2),
            _ret(self.SHIP, "TK10000282", self.SKU_H1, "00001", 6, row=3),
            _ret(self.SHIP, "TK10000282", self.SKU_H1, "00002", 6, row=4),
            *[
                _ret(self.SHIP, "TK10000282", self.SKU_H1, f"0000{i}", 4, row=i + 1)
                for i in range(3, 9)
            ],
        ]
        rows_h2 = [
            _ret(self.SHIP, "TK10000282", self.SKU_H2, "00009", 6, row=11),
            _ret(self.SHIP, "TK10000282", self.SKU_H2, "00010", 3, row=12),
        ]
        inv = [
            _inv(self.SHIP, self.SKU_H1, "ZP", 23, row=2),
            _inv(self.SHIP, self.SKU_H1, "QC", 13, row=3),
            _inv(self.SHIP, self.SKU_H1, "NG", 6, row=4),
            _inv(self.SHIP, self.SKU_H2, "ZP", 9, row=5),
        ]
        return rows_h1, rows_h2, inv

    def test_e10_and_cascade_rollback(self):
        rows_h1, rows_h2, inv = self._data()
        result, ledger = _run(
            self.SHIP, [self.SKU_H1, self.SKU_H2],
            {self.SKU_H1: rows_h1, self.SKU_H2: rows_h2}, inv,
            _cfg(max_backtrack=10),
        )
        assert len(result.group_results) == 2

        g1 = result.group_results[0]
        assert g1.success is False
        assert g1.error_code == ERR_E10
        assert g1.stats.backtrack_count == 11  # 第 11 次回溯（>上限10）触发 E10
        assert g1.details == []

        g2 = result.group_results[1]
        assert g2.success is False
        assert g2.error_code == ERROR_CASCADE_ROLLBACK
        assert g2.stats.backtrack_count == 0
        assert g2.stats.precheck_hit == ""
        assert g2.details == []

        # E10 全量回滚后 H1 账本还原；H2 从未执行，库存原封未动
        assert ledger.query_qc_total(self.SHIP, self.SKU_H1, "ZP") == 23
        assert ledger.query_qc_total(self.SHIP, self.SKU_H1, "QC") == 13
        assert ledger.query_qc_total(self.SHIP, self.SKU_H1, "NG") == 6
        assert ledger.query_qc_total(self.SHIP, self.SKU_H2, "ZP") == 9

    def test_cascade_events_simple(self):
        """简版：H2 连带回滚事件子类型=连带回滚—跨SKU短路，处理序/候选填 '-'。"""
        rows_h1, rows_h2, inv = self._data()
        result, _ = _run(
            self.SHIP, [self.SKU_H1, self.SKU_H2],
            {self.SKU_H1: rows_h1, self.SKU_H2: rows_h2}, inv,
            _cfg(max_backtrack=10, level=DEBUG_LEVEL_SIMPLE),
        )
        g1, g2 = result.group_results
        assert len(g1.events) == 9
        assert len(g2.events) == 2
        for e in g2.events:
            assert e.error_code == ERROR_CASCADE_ROLLBACK
            assert e.fail_sub_type == FAIL_SUB_CASCADE_CROSS_SKU
            assert e.process_order == "-"
            assert e.candidate_qc_count == "-"
            assert e.backtrack_no == 0


class TestPrecheckShortCircuit:
    """预检测 A/B 命中 → 整组 E09 + precheck_hit 记录 + 同物流单号短路（§4.2.3）。"""

    def test_precheck_a_hit_and_short_circuit(self):
        """预检测A：行 D=3 的初始可用QC数=0（ZP T=10 ≠3 且 < 3+10）→ E09 + 短路。"""
        ship, sku_a, sku_b = "SF010", "SKU-A", "SKU-B"
        rows_a = [
            _ret(ship, "WMS1", sku_a, "00001", 3),
            _ret(ship, "WMS1", sku_a, "00002", 10),
        ]
        rows_b = [_ret(ship, "WMS1", sku_b, "00001", 1)]
        inv = [
            _inv(ship, sku_a, "ZP", 10),
            _inv(ship, sku_b, "NG", 1, lot_no="LB01"),
        ]
        result, ledger = _run(
            ship, [sku_a, sku_b], {sku_a: rows_a, sku_b: rows_b}, inv,
            _cfg(level=DEBUG_LEVEL_SIMPLE),
        )
        g1, g2 = result.group_results
        assert g1.error_code == ERR_E09
        assert g1.stats.precheck_hit == PRECHECK_HIT_A
        assert g1.stats.backtrack_count == 0
        assert g2.error_code == ERROR_CASCADE_ROLLBACK
        # 预检测命中的组不进入分配循环：两个 SKU 的库存均未动
        assert ledger.query_qc_total(ship, sku_a, "ZP") == 10
        assert ledger.query_qc_total(ship, sku_b, "NG") == 1
        # 失败子类型：初始可用QC=0 的行与子类型文案一致（第 19 列）
        subtypes = {e.line_no: e.fail_sub_type for e in g1.events}
        assert subtypes["00001"] == FAIL_SUB_PRECHECK_A
        assert subtypes["00002"] == FAIL_SUB_PRECHECK_A
        assert all(e.error_code == ERR_E09 for e in g1.events)

    def test_precheck_b_hit(self):
        """预检测B（spec §4.2.3 示例）：需求 {1,2,5,6}，库存 ZP=7/QC=4/NG=3，
        D=5 与 D=6 强制竞争 ZP，S=11 > T=7 且碎片无法消化 → E09。"""
        ship, sku = "SF011", "SKU-PB"
        rows = [
            _ret(ship, "WMS1", sku, "00001", 1),
            _ret(ship, "WMS1", sku, "00002", 2),
            _ret(ship, "WMS1", sku, "00003", 5),
            _ret(ship, "WMS1", sku, "00004", 6),
        ]
        inv = [
            _inv(ship, sku, "ZP", 7),
            _inv(ship, sku, "QC", 4),
            _inv(ship, sku, "NG", 3),
        ]
        result, ledger = _run(ship, [sku], {sku: rows}, inv, _cfg(level=DEBUG_LEVEL_SIMPLE))
        group = result.group_results[0]
        assert group.error_code == ERR_E09
        assert group.stats.precheck_hit == PRECHECK_HIT_B
        assert group.details == []
        assert all(e.fail_sub_type == FAIL_SUB_PRECHECK_B for e in group.events)
        assert ledger.query_qc_total(ship, sku, "ZP") == 7


class TestLineStatus:
    """行状态按实际(批号+效期)组合数判定（§4.2.4 步骤 3 / §4.4.1）。"""

    def test_strategy_three_multi_lot_is_manual(self):
        """策略三跨批号拼凑（9+3=12）→ 2 种组合 → 手工操作，多条明细。"""
        ship, sku = "SF020", "SKU-M"
        rows = [_ret(ship, "WMS1", sku, "00001", 12)]
        inv = [
            _inv(ship, sku, "ZP", 9, lot_no="L01"),
            _inv(ship, sku, "ZP", 3, lot_no="W01"),
        ]
        result, _ = _run(ship, [sku], {sku: rows}, inv)
        group = result.group_results[0]
        assert group.success is True
        assert len(group.details) == 2
        assert sum(d.alloc_qty for d in group.details) == 12
        assert all(d.strategy_used == STRATEGY_THREE for d in group.details)
        assert all(d.line_status == STATUS_MANUAL for d in group.details)

    def test_strategy_one_two_are_batch_import(self):
        """策略一/二必然单五元组 → 批量导入（推论验证，见 §4.4.1）。"""
        ship, sku = "SF021", "SKU-N"
        rows = [
            _ret(ship, "WMS1", sku, "00001", 5),
            _ret(ship, "WMS1", sku, "00002", 3),
        ]
        inv = [_inv(ship, sku, "ZP", 8)]
        result, _ = _run(ship, [sku], {sku: rows}, inv)
        group = result.group_results[0]
        assert group.success is True
        assert {d.strategy_used for d in group.details} == {STRATEGY_ONE, STRATEGY_TWO}
        assert all(d.line_status == STATUS_BATCH_IMPORT for d in group.details)


class TestEdgeCases:
    """边界：空 SKU 列表、缺失静态计划。"""

    def test_empty_sku_list(self):
        result, _ = _run("SF030", [], {}, [])
        assert result.shipment_no == "SF030"
        assert result.group_results == []

    def test_missing_plan_is_e09(self):
        """plan_map 中找不到计划 → 该组 E09（对应 VBA 传空 Dictionary）。"""
        ship, sku = "SF031", "SKU-X"
        ledger = build_ledger([_inv(ship, sku, "ZP", 1)])
        result = allocate_shipment(ship, [sku], {}, {}, ledger, _cfg())
        group = result.group_results[0]
        assert group.success is False
        assert group.error_code == ERR_E09
