"""M07 排序·预检测·QC筛选单元测试。

对应 VBA modTestRunner.bas 的 TC-SF01~SF09，以及需求 TC-06（预检测A）、
TC-11（静态排序四级规则）、TC-12（动态筛选与 nextMinQty）、TC-40（预检测B，R082）。
"""

from returned_inventory.ledger import build_ledger, new_undo_log
from returned_inventory.models import (
    InventoryKey,
    NormalizedInventoryLine,
    NormalizedReturnLine,
)
from returned_inventory.sort_filter import (
    StaticPlan,
    build_static_plan,
    filter_candidate_pool,
    make_plan_line_key,
    run_precheck,
)

EXPIRY = "2029/01/01"


def _ret_line(
    shipment_no="SF001",
    wms_order_no="WMS1",
    sku="SKU-A",
    line_no="00001",
    qty=1,
    excel_row_num=2,
):
    return NormalizedReturnLine(
        excel_row_num=excel_row_num,
        shipment_no=shipment_no,
        wms_order_no=wms_order_no,
        sku=sku,
        line_no=line_no,
        qty=qty,
        line_no_valid=True,
        qty_valid=True,
        empty_fields="",
    )


def _inv_line(
    shipment_no="SF001",
    sku="SKU-A",
    qc="ZP",
    lot_no="L01",
    expiry=EXPIRY,
    qty=10,
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
        qc_valid=True,
        expiry_valid=True,
        qty_valid=True,
        empty_fields="",
    )


def _tc12_ledger(ship="SF001", sku="SKU-A"):
    """TC-12 库存：ZP:2, QC:5, NG:5。"""
    return build_ledger(
        [
            _inv_line(ship, sku, "ZP", qty=2),
            _inv_line(ship, sku, "QC", qty=5),
            _inv_line(ship, sku, "NG", qty=5),
        ]
    )


def _tc12_rows(ship="SF001", sku="SKU-A"):
    """TC-12 退单行：D={2,5,5}。"""
    return [
        _ret_line(ship, "WMS_TC12", sku, "00001", 2),
        _ret_line(ship, "WMS_TC12", sku, "00002", 5),
        _ret_line(ship, "WMS_TC12", sku, "00003", 5),
    ]


class TestBuildStaticPlanSort:
    """静态排序四级规则（TC-11，对应 VBA TC-SF01/TC-SF02）。"""

    def test_tc12_init_qc_count_asc(self):
        """TC-SF01：可用QC数升序为主排序；同级行号升序兜底。

        行00001（D=2, nextMinQty=5）：ZP(T=2=D→可用)，QC/NG(T=5<2+5=7→不可用) → 1
        行00002/00003（D=5, nextMinQty=2）：QC/NG(T=5=D→可用)，ZP(T=2<7→不可用) → 2
        """
        plan = build_static_plan(_tc12_rows(), _tc12_ledger())
        assert plan.row_count == 3
        assert plan.group_min_qty == 2
        assert [l.line_no for l in plan.lines] == ["00001", "00002", "00003"]
        assert [l.init_qc_count for l in plan.lines] == [1, 2, 2]

    def test_qty_desc_tiebreak(self):
        """TC-SF02：可用QC数相同时 Qty 降序（第二级）。"""
        ledger = build_ledger(
            [_inv_line(qc="ZP", qty=20), _inv_line(qc="QC", qty=20)]
        )
        rows = [
            _ret_line(line_no="00001", qty=4),  # 输入顺序小需求在前
            _ret_line(line_no="00002", qty=8),
        ]
        plan = build_static_plan(rows, ledger)
        assert [l.qty for l in plan.lines] == [8, 4]
        assert [l.line_no for l in plan.lines] == ["00002", "00001"]
        assert [l.init_qc_count for l in plan.lines] == [2, 2]

    def test_four_level_sort_all_triggered(self):
        """TC-11：四级规则全触发——打乱输入顺序，验证完整排序链。

        库存 ZP:9, QC:20, NG:20：
        - R1 D=9：ZP(9=D ✓)、QC/NG(20≥9+5=14 ✓) → 可用QC数 3
        - R2 D=8：ZP(9≠8 且 9<8+5=13 ✗)、QC/NG(20≥13 ✓) → 2
        - R3~R6 D=5：ZP(9≠5 且 9<10 ✗)、QC/NG(20≥10 ✓) → 2
        预期顺序（一级 可用QC数升序 → 二级 Qty降序 → 三级 WMS升序 → 四级 行号升序）：
        R2(8) → R6(W01:00001) → R5(W01:00002) → R4(W02:00001) → R3(W03:00001) → R1(9)
        """
        ledger = build_ledger(
            [
                _inv_line(qc="ZP", qty=9),
                _inv_line(qc="QC", qty=20),
                _inv_line(qc="NG", qty=20),
            ]
        )
        rows = [
            _ret_line(wms_order_no="W09", line_no="00001", qty=9),   # R1
            _ret_line(wms_order_no="W03", line_no="00001", qty=5),   # R3
            _ret_line(wms_order_no="W01", line_no="00002", qty=5),   # R5
            _ret_line(wms_order_no="W00", line_no="00001", qty=8),   # R2
            _ret_line(wms_order_no="W01", line_no="00001", qty=5),   # R6
            _ret_line(wms_order_no="W02", line_no="00001", qty=5),   # R4
        ]
        plan = build_static_plan(rows, ledger)
        assert [(l.wms_order_no, l.line_no) for l in plan.lines] == [
            ("W00", "00001"),  # R2：可用QC数=2 中 Qty 最大
            ("W01", "00001"),  # R6：W01 内行号升序
            ("W01", "00002"),  # R5
            ("W02", "00001"),  # R4：WMS退单号升序
            ("W03", "00001"),  # R3
            ("W09", "00001"),  # R1：可用QC数=3 排最后
        ]
        assert [l.init_qc_count for l in plan.lines] == [2, 2, 2, 2, 2, 3]

    def test_string_compare_no_numeric_conversion(self):
        """WMS退单号/行号按字符串字典序比较，不做数值转换（需求 §4.2.1）。"""
        ledger = build_ledger([_inv_line(qc="ZP", qty=100)])
        rows = [
            _ret_line(wms_order_no="WMS10", line_no="00001", qty=5),
            _ret_line(wms_order_no="WMS9", line_no="00001", qty=5),
        ]
        plan = build_static_plan(rows, ledger)
        # 字符串升序："WMS10" < "WMS9"（'1' < '9'）；数值比较则相反
        assert [l.wms_order_no for l in plan.lines] == ["WMS10", "WMS9"]

    def test_empty_input_returns_empty_plan(self):
        """空输入返回空计划，不抛错（对应 VBA 返回不含 RowCount 的空 Dictionary）。"""
        plan = build_static_plan([], _tc12_ledger())
        assert plan == StaticPlan()
        assert plan.row_count == 0


class TestInitQCCount:
    """初始可用QC数计算：需求 §4.2.3 示例（需求 1,2,5,6 + ZP-7/QC-4/NG-3）。"""

    def test_spec_423_example(self):
        ledger = build_ledger(
            [
                _inv_line(qc="ZP", qty=7),
                _inv_line(qc="QC", qty=4),
                _inv_line(qc="NG", qty=3),
            ]
        )
        rows = [_ret_line(line_no=f"0000{i}", qty=q) for i, q in enumerate([1, 2, 5, 6], 1)]
        plan = build_static_plan(rows, ledger)

        # 每行 nextMinQty = 同组其他行需求最小值：
        # D=6: nmq=1 → ZP(7≥7 ✓) → 1；D=5: nmq=1 → ZP(7≥6 ✓) → 1
        # D=2: nmq=1 → ZP/QC/NG 均 ≥3 ✓ → 3；D=1: nmq=2 → 均 ≥3 ✓ → 3
        count_by_qty = {l.qty: l.init_qc_count for l in plan.lines}
        assert count_by_qty == {6: 1, 5: 1, 2: 3, 1: 3}
        # 排序结果：可用QC数升序 → Qty降序
        assert [l.qty for l in plan.lines] == [6, 5, 2, 1]

    def test_single_row_is_last_row_mode(self):
        """整组只有 1 行时按"最后行"处理：仅 T=D 可用（nextMinQty 不参与）。"""
        ledger = build_ledger(
            [_inv_line(qc="ZP", qty=5), _inv_line(qc="QC", qty=9)]
        )
        plan = build_static_plan([_ret_line(qty=5)], ledger)
        # ZP(5=D ✓)、QC(9≠5 ✗，最后行不允许 T>D) → 1
        assert [l.init_qc_count for l in plan.lines] == [1]


class TestRunPrecheck:
    """预检测A/B（TC-06 / TC-40，对应 VBA TC-SF03~SF05）。"""

    def test_precheck_a_hit(self):
        """TC-SF03 / TC-06：库存 ZP:5, QC:6，需求 {10,1}。

        D=10 行：ZP(5≠10 且 5<11 ✗)、QC(6≠10 且 6<11 ✗) → 可用QC数=0 → 预检测A命中。
        """
        ledger = build_ledger([_inv_line(qc="ZP", qty=5), _inv_line(qc="QC", qty=6)])
        rows = [_ret_line(line_no="00001", qty=10), _ret_line(line_no="00002", qty=1)]
        plan = build_static_plan(rows, ledger)
        assert [l.init_qc_count for l in plan.lines] == [0, 0]

        result = run_precheck(plan, ledger)
        assert result.precheck_a_hit is True
        assert result.precheck_b_hit is False  # A 命中后不再判断 B

    def test_precheck_neither_hit(self):
        """TC-SF04：TC-12 正常数据（init_qc_count={1,2,2}，无 =0，无 B 竞争）。"""
        plan = build_static_plan(_tc12_rows(), _tc12_ledger())
        result = run_precheck(plan, _tc12_ledger())
        assert result.precheck_a_hit is False
        assert result.precheck_b_hit is False

    def test_precheck_b_hit_spec_example(self):
        """TC-40 / §4.2.3 示例：需求 {5,6,1}，库存 ZP-7/QC-4/NG-1。

        D=5、D=6 均仅可用 ZP（强制竞争），S=11，T=7；
        非强制行 D=1（也可用 ZP）提供 minQtyOther=1。
        T≠S 且 T=7 < S+minQtyOther=12 → 预检测B命中。
        """
        ledger = build_ledger(
            [
                _inv_line(qc="ZP", qty=7),
                _inv_line(qc="QC", qty=4),
                _inv_line(qc="NG", qty=1),
            ]
        )
        rows = [
            _ret_line(line_no="00001", qty=5),
            _ret_line(line_no="00002", qty=6),
            _ret_line(line_no="00003", qty=1),
        ]
        plan = build_static_plan(rows, ledger)
        count_by_qty = {l.qty: l.init_qc_count for l in plan.lines}
        assert count_by_qty == {5: 1, 6: 1, 1: 2}

        result = run_precheck(plan, ledger)
        assert result.precheck_a_hit is False
        assert result.precheck_b_hit is True

    def test_precheck_b_hit_no_non_forced_row(self):
        """TC-SF05：无非强制行且 T≠S → 命中（minQtyOther 视为 +∞）。

        库存仅 ZP:3，两行 D 均=3：均锁定 ZP（T=3=D 可用），S=6 > T=3。
        """
        ledger = build_ledger([_inv_line(qc="ZP", qty=3)])
        rows = [_ret_line(line_no="00001", qty=3), _ret_line(line_no="00002", qty=3)]
        plan = build_static_plan(rows, ledger)
        assert [l.init_qc_count for l in plan.lines] == [1, 1]

        result = run_precheck(plan, ledger)
        assert result.precheck_a_hit is False
        assert result.precheck_b_hit is True

    def test_precheck_b_not_hit_when_t_equals_s(self):
        """不命中条件一：T=S（总量恰好等于合计需求，精确可行）。"""
        ledger = build_ledger([_inv_line(qc="ZP", qty=6)])
        rows = [_ret_line(line_no="00001", qty=3), _ret_line(line_no="00002", qty=3)]
        plan = build_static_plan(rows, ledger)
        assert [l.init_qc_count for l in plan.lines] == [1, 1]

        result = run_precheck(plan, ledger)
        assert result.precheck_a_hit is False
        assert result.precheck_b_hit is False  # T=6 = S=6

    def test_precheck_b_not_hit_when_t_covers_min_qty_other(self):
        """不命中条件二：T ≥ S + minQtyOther（碎片能被非强制行消化）。

        需求 {3,3,4}，库存 ZP:10、QC:4：
        D=3 两行锁定 ZP（10≥3+3=6 ✓，QC 4<6 ✗），S=6；
        D=4 非强制（ZP 10≥4+3=7 ✓、QC 4=D ✓），minQtyOther=4。
        T=10 ≥ S+minQtyOther=10 → 不命中。
        """
        ledger = build_ledger([_inv_line(qc="ZP", qty=10), _inv_line(qc="QC", qty=4)])
        rows = [
            _ret_line(line_no="00001", qty=3),
            _ret_line(line_no="00002", qty=3),
            _ret_line(line_no="00003", qty=4),
        ]
        plan = build_static_plan(rows, ledger)
        count_by_qty = {l.qty: l.init_qc_count for l in plan.lines}
        assert count_by_qty == {3: 1, 4: 2}

        result = run_precheck(plan, ledger)
        assert result.precheck_a_hit is False
        assert result.precheck_b_hit is False

    def test_precheck_b_hit_fragment_case(self):
        """S < T < S + minQtyOther 的碎片场景：按需求 §4.2.3 完整判定命中。

        需求 {3,3,4}，库存 ZP:8、QC:4：S=6，T=8，minQtyOther=4，
        T≠S 且 8 < 10 → 命中（剩余碎片 2 无任何行能消化）。
        注：VBA 代码此处实现为 supply < forcedDemand（不命中），
        本移植按需求 §4.2.3 / §6.3.1.2 伪代码实现，属有意偏差。
        """
        ledger = build_ledger([_inv_line(qc="ZP", qty=8), _inv_line(qc="QC", qty=4)])
        rows = [
            _ret_line(line_no="00001", qty=3),
            _ret_line(line_no="00002", qty=3),
            _ret_line(line_no="00003", qty=4),
        ]
        plan = build_static_plan(rows, ledger)
        result = run_precheck(plan, ledger)
        assert result.precheck_a_hit is False
        assert result.precheck_b_hit is True

    def test_precheck_read_only(self):
        """run_precheck 只读：plan 与 ledger 均不被修改。"""
        ledger = _tc12_ledger()
        plan = build_static_plan(_tc12_rows(), ledger)
        keys_before = {
            k: (ledger.get_original_qty(k), ledger.get_current_qty(k))
            for k in ledger.get_ledger_keys()
        }
        run_precheck(plan, ledger)
        assert plan == build_static_plan(_tc12_rows(), _tc12_ledger())
        assert {
            k: (ledger.get_original_qty(k), ledger.get_current_qty(k))
            for k in ledger.get_ledger_keys()
        } == keys_before

    def test_precheck_empty_plan(self):
        """空/None 计划防御：返回默认值（不命中），不抛错。"""
        ledger = _tc12_ledger()
        assert run_precheck(None, ledger) == run_precheck(StaticPlan(), ledger)
        assert run_precheck(StaticPlan(), ledger).precheck_a_hit is False


class TestFilterCandidatePool:
    """动态筛选（TC-12，对应 VBA TC-SF06~SF09）。"""

    def test_middle_row_dynamic_next_min_qty(self):
        """TC-SF06：首行筛选——动态 nextMinQty=5（而非 groupMinQty=2）。

        当前行 00001（D=2），后续行 {5,5} → nextMinQty=5：
        ZP(T=2=D ✓)；QC/NG(T=5≠2 且 5<2+5=7 ✗) → 候选池仅 ZP 一行。
        """
        ledger = _tc12_ledger()
        plan = build_static_plan(_tc12_rows(), ledger)
        pool = filter_candidate_pool(make_plan_line_key("WMS_TC12", "00001"), plan, ledger)
        assert len(pool) == 1
        assert pool[0].qc == "ZP"
        assert pool[0].current_qty == 2

    def test_next_min_qty_changes_with_position(self):
        """动态 nextMinQty 随处理进度变化：排序位置越靠后，剩余行集合越小。

        当前行 00002（D=5，位置2），后续行 {5} → nextMinQty=5：
        ZP(T=2 ✗)；QC/NG(T=5=D ✓) → 候选池 QC+NG 两行。
        """
        ledger = _tc12_ledger()
        plan = build_static_plan(_tc12_rows(), ledger)
        pool = filter_candidate_pool(make_plan_line_key("WMS_TC12", "00002"), plan, ledger)
        assert sorted(r.qc for r in pool) == ["NG", "QC"]

    def test_last_row_requires_exact_match(self):
        """TC-SF07：最后行仅 T=D 有效。

        单行 D=5，库存 QC:5、NG:5、ZP:0：QC/NG(T=5=D ✓)，ZP(T=0 不参与)。
        """
        ledger = build_ledger([_inv_line(qc="QC", qty=5), _inv_line(qc="NG", qty=5)])
        plan = build_static_plan([_ret_line(qty=5)], ledger)
        pool = filter_candidate_pool(make_plan_line_key("WMS1", "00001"), plan, ledger)
        assert len(pool) == 2
        assert sorted(r.qc for r in pool) == ["NG", "QC"]

    def test_last_row_rejects_t_greater_than_d(self):
        """最后行 T>D 不可用（库存必须精确耗尽，不能留碎片）。

        单行 D=5（整组一行即最后行）：ZP(T=5=D ✓)、QC(T=7≠5 ✗，最后行不允许 T>D)。
        """
        ledger = build_ledger([_inv_line(qc="ZP", qty=5), _inv_line(qc="QC", qty=7)])
        plan = build_static_plan([_ret_line(qty=5)], ledger)
        pool = filter_candidate_pool(make_plan_line_key("WMS1", "00001"), plan, ledger)
        assert len(pool) == 1
        assert pool[0].qc == "ZP"

    def test_tried_qcs_excluded(self):
        """TC-SF08：tried_qcs 中的 QC 被排除（回溯时本行已尝试失败）。"""
        ledger = build_ledger([_inv_line(qc="ZP", qty=5), _inv_line(qc="QC", qty=5)])
        plan = build_static_plan([_ret_line(qty=5)], ledger)
        key = make_plan_line_key("WMS1", "00001")

        pool = filter_candidate_pool(key, plan, ledger, ["QC"])
        assert len(pool) == 1
        assert pool[0].qc == "ZP"

        # 全部 QC 均已尝试 → 空候选池
        assert filter_candidate_pool(key, plan, ledger, ["ZP", "QC"]) == []

    def test_zero_current_qty_rows_filtered(self):
        """候选池过滤 current_qty=0 的五元组行（同一 QC 下的已耗尽批号）。"""
        ledger = build_ledger(
            [
                _inv_line(qc="ZP", lot_no="L01", qty=2),
                _inv_line(qc="ZP", lot_no="L02", qty=3),
            ]
        )
        # L01 被前序行耗尽：T 从 5 降为 3
        undo_log = new_undo_log()
        assert ledger.deduct(InventoryKey("SF001", "SKU-A", "ZP", "L01", EXPIRY), 2, undo_log)

        plan = build_static_plan([_ret_line(qty=3)], ledger)
        pool = filter_candidate_pool(make_plan_line_key("WMS1", "00001"), plan, ledger)
        assert len(pool) == 1
        assert pool[0].lot_no == "L02"
        assert pool[0].current_qty == 3

    def test_unknown_line_key_returns_empty(self):
        """行键在 plan 中找不到（调用参数有误）→ 安全返回空。"""
        ledger = _tc12_ledger()
        plan = build_static_plan(_tc12_rows(), ledger)
        assert filter_candidate_pool("NOPE:99999", plan, ledger) == []

    def test_empty_plan_returns_empty(self):
        ledger = _tc12_ledger()
        assert filter_candidate_pool("WMS:00001", None, ledger) == []
        assert filter_candidate_pool("WMS:00001", StaticPlan(), ledger) == []

    def test_same_next_min_qty_definition_as_static(self):
        """TC-SF09：首行动态筛选与静态计算使用完全相同的 nextMinQty 定义。

        账本未扣减时，首行 init_qc_count 与候选池中的 QC 种类数一致。
        """
        ledger = _tc12_ledger()
        plan = build_static_plan(_tc12_rows(), ledger)
        pool = filter_candidate_pool(make_plan_line_key("WMS_TC12", "00001"), plan, ledger)
        dynamic_qc_count = len({r.qc for r in pool})
        assert plan.lines[0].init_qc_count == 1
        assert dynamic_qc_count == plan.lines[0].init_qc_count

    def test_filter_uses_live_ledger_not_plan_cache(self):
        """filter 结果不受 plan 静态缓存影响：T 取账本实时值。

        TC-12 首行分配后（ZP 扣减 2），对第二行筛选：ZP(T=0 ✗)，
        候选池只剩 QC+NG——若误用 plan 缓存的静态可用性会错误包含 ZP。
        """
        ledger = _tc12_ledger()
        plan = build_static_plan(_tc12_rows(), ledger)
        undo_log = new_undo_log()
        assert ledger.deduct(InventoryKey("SF001", "SKU-A", "ZP", "L01", EXPIRY), 2, undo_log)

        pool = filter_candidate_pool(make_plan_line_key("WMS_TC12", "00002"), plan, ledger)
        assert sorted(r.qc for r in pool) == ["NG", "QC"]
        # plan 本身不被 filter 修改
        assert [l.init_qc_count for l in plan.lines] == [1, 2, 2]
