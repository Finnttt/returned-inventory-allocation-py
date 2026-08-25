"""M11 状态判定单元测试（对应 VBA modStatus.bas / 需求 §4.4、§5.1）。

覆盖 TC-04（全批量导入聚合）、TC-05（手工操作传染聚合）、
TC-28（同退单号多错误码升序分号合并）、TC-37（多物流单号多错误码并存），
以及整单回滚（校验阶段 E08 / 分配阶段 E09、E10）、连带回滚原因格式、
E07 孤立物流单号 [N/A] 专用记录。

测试数据全部内存构造，不接触 Excel。
"""

from returned_inventory.models import (
    ERR_E01,
    ERR_E02,
    ERR_E03,
    ERR_E04,
    ERR_E08,
    ERR_E09,
    ERR_E10,
    ERROR_CASCADE_ROLLBACK,
    LINE_STATUS_FAILED,
    NA_PLACEHOLDER,
    QC_ZP,
    SOURCE_INVENTORY_TABLE,
    SOURCE_RETURN_TABLE,
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STATUS_UNALLOCATED,
    AllocationDetail,
    GroupAllocResult,
    GroupStats,
    NormalizedReturnLine,
    ShipmentAllocResult,
    ValidationIssue,
    ValidationResult,
)
from returned_inventory.status import (
    FinalResult,
    aggregate_wms_status,
    apply_rollback,
    build_rollback_reason,
    determine_line_status,
)


def make_detail(
    shipment_no="SH1",
    wms_order_no="W1",
    sku="SKU1",
    line_no="00001",
    order_qty=10,
    qc=QC_ZP,
    lot_no="LA01",
    expiry="2029/01/01",
    alloc_qty=10,
    line_status="",
    strategy_used="",
):
    return AllocationDetail(
        shipment_no=shipment_no,
        wms_order_no=wms_order_no,
        sku=sku,
        line_no=line_no,
        order_qty=order_qty,
        qc=qc,
        lot_no=lot_no,
        expiry=expiry,
        alloc_qty=alloc_qty,
        line_status=line_status,
        strategy_used=strategy_used,
    )


def make_group(shipment_no, sku, success, error_code="", details=None):
    return GroupAllocResult(
        shipment_no=shipment_no,
        sku=sku,
        success=success,
        error_code=error_code,
        stats=GroupStats(shipment_no, sku, 0, ""),
        details=details or [],
    )


def make_shipment(shipment_no, groups):
    return ShipmentAllocResult(shipment_no=shipment_no, group_results=groups)


def make_order(
    excel_row_num=2,
    shipment_no="SH1",
    wms_order_no="W1",
    sku="SKU1",
    line_no="00001",
    qty=10,
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


def make_issue(
    shipment_no="SH1",
    wms_order_no="W1",
    sku="SKU1",
    error_code=ERR_E01,
    source_table=SOURCE_RETURN_TABLE,
    excel_row_num=2,
    reason="",
):
    return ValidationIssue(
        shipment_no=shipment_no,
        wms_order_no=wms_order_no,
        sku=sku,
        error_code=error_code,
        source_table=source_table,
        excel_row_num=excel_row_num,
        field_name="",
        raw_value="",
        reason=reason,
    )


EMPTY_VALIDATION = ValidationResult(has_failures=False, failed_shipment_count=0)


# -----------------------------------------------------------------------------
# determine_line_status：行级状态判定（§4.4.1）
# -----------------------------------------------------------------------------


class TestDetermineLineStatus:
    def test_single_combo_is_batch_import(self):
        details = [make_detail(line_no="00001", lot_no="LA01", expiry="2029/01/01")]
        assert determine_line_status(details, "00001") == STATUS_BATCH_IMPORT

    def test_two_combos_is_manual(self):
        details = [
            make_detail(line_no="00001", lot_no="LA03", expiry="2031/01/01", alloc_qty=1),
            make_detail(line_no="00001", lot_no="LA02", expiry="2030/01/01", alloc_qty=1),
        ]
        assert determine_line_status(details, "00001") == STATUS_MANUAL

    def test_same_combo_repeated_still_batch(self):
        # 同一批号+效期出现多条明细，组合数仍为 1
        details = [
            make_detail(line_no="00001", lot_no="LA01", alloc_qty=6),
            make_detail(line_no="00001", lot_no="LA01", alloc_qty=4),
        ]
        assert determine_line_status(details, "00001") == STATUS_BATCH_IMPORT

    def test_no_matching_detail_is_failed(self):
        details = [make_detail(line_no="00002")]
        assert determine_line_status(details, "00001") == LINE_STATUS_FAILED

    def test_zero_alloc_qty_not_counted(self):
        details = [make_detail(line_no="00001", alloc_qty=0)]
        assert determine_line_status(details, "00001") == LINE_STATUS_FAILED

    def test_empty_details_is_failed(self):
        assert determine_line_status([], "00001") == LINE_STATUS_FAILED


# -----------------------------------------------------------------------------
# build_rollback_reason：原因字符串格式（§5.1 R121）
# -----------------------------------------------------------------------------


class TestBuildRollbackReason:
    def test_cascade_format(self):
        assert build_rollback_reason([], "E10") == "整单回滚（触发原因：E10）"

    def test_single_direct_code(self):
        assert build_rollback_reason(["E10"], "E10") == "E10 - 回溯超限"

    def test_multi_codes_sorted_ascending(self):
        # 乱序输入 → 按 E01→E99 升序输出
        reason = build_rollback_reason(["E10", "E08", "E09"], "E08")
        assert reason == (
            "E08 - 同物流单号+SKU数量不一致; "
            "E09 - 分配路径穷尽; "
            "E10 - 回溯超限"
        )

    def test_duplicate_codes_deduped(self):
        reason = build_rollback_reason(["E04", "E04", "E01"], "E01")
        assert reason == "E01 - 关键字段为空或格式异常; E04 - 数量非法"

    def test_unknown_code_fallback_text(self):
        assert build_rollback_reason(["E42"], "E42") == "E42 - 未知错误"


# -----------------------------------------------------------------------------
# apply_rollback：校验阶段失败 → 整单回滚
# -----------------------------------------------------------------------------


class TestValidationStageRollback:
    def test_tc37_two_shipments_independent_errors(self):
        """TC-37：SF0037 命中 E01+E02+E04，SF0038 命中 E08，两组互不干扰。"""
        s37 = "SF3190000000037"
        s38 = "SF3190000000038"
        orders = [
            make_order(excel_row_num=2, shipment_no=s37, wms_order_no="TK10000370",
                       sku="H000000037", line_no="X1234", qty=5),
            make_order(excel_row_num=3, shipment_no=s37, wms_order_no="TK10000370",
                       sku="H000000037", line_no="00002", qty=-2),
            make_order(excel_row_num=4, shipment_no=s38, wms_order_no="TK10000380",
                       sku="H000000038", line_no="00001", qty=6),
            make_order(excel_row_num=5, shipment_no=s38, wms_order_no="TK10000381",
                       sku="H000000038", line_no="00001", qty=4),
        ]
        issues = [
            make_issue(shipment_no=s37, wms_order_no="TK10000370",
                       sku="H000000037", error_code=ERR_E01, excel_row_num=2),
            make_issue(shipment_no=s37, wms_order_no="TK10000370",
                       sku="H000000037", error_code=ERR_E02, excel_row_num=3),
            make_issue(shipment_no=s37, wms_order_no="TK10000370",
                       sku="H000000037", error_code=ERR_E04, excel_row_num=3),
            # E08 挂在物流单号+SKU 级（WMS 为 [N/A]），扩展到含该 SKU 的所有退单号
            make_issue(shipment_no=s38, wms_order_no=NA_PLACEHOLDER,
                       sku="H000000038", error_code=ERR_E08,
                       source_table=NA_PLACEHOLDER, excel_row_num=0),
        ]
        validation = ValidationResult(has_failures=True, failed_shipment_count=2, issues=issues)

        result = apply_rollback([], validation, issues, orders)

        assert result.details == []
        entries = result.summary_entries
        assert len(entries) == 3
        assert (entries[0].shipment_no, entries[0].wms_order_no) == (s37, "TK10000370")
        assert entries[0].status == STATUS_UNALLOCATED
        assert entries[0].reason == (
            "E01 - 关键字段为空或格式异常; "
            "E02 - 退单表行号重复或不连续; "
            "E04 - 数量非法"
        )
        # E08 对该 SKU 下的两个退单号均为直接原因
        assert (entries[1].wms_order_no, entries[1].reason) == (
            "TK10000380", "E08 - 同物流单号+SKU数量不一致"
        )
        assert (entries[2].wms_order_no, entries[2].reason) == (
            "TK10000381", "E08 - 同物流单号+SKU数量不一致"
        )

    def test_tc28_four_codes_merged_ascending(self):
        """TC-28：同一退单号命中 E01+E02+E03+E04（E03 来自库存表），升序分号合并。"""
        s64 = "SF3190000000064"
        orders = [
            make_order(excel_row_num=2, shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", line_no="X1234", qty=5),
            make_order(excel_row_num=3, shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", line_no="00002", qty=-3),
            make_order(excel_row_num=4, shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", line_no="00003", qty=2),
        ]
        # 故意乱序 + 重复 E04，验证排序与同码去重
        issues = [
            make_issue(shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", error_code=ERR_E04, excel_row_num=3),
            make_issue(shipment_no=s64, wms_order_no=NA_PLACEHOLDER,
                       sku="H000000064", error_code=ERR_E03,
                       source_table=SOURCE_INVENTORY_TABLE, excel_row_num=2),
            make_issue(shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", error_code=ERR_E02, excel_row_num=3),
            make_issue(shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", error_code=ERR_E01, excel_row_num=2),
            make_issue(shipment_no=s64, wms_order_no="TK10000640",
                       sku="H000000064", error_code=ERR_E04, excel_row_num=3),
        ]
        validation = ValidationResult(has_failures=True, failed_shipment_count=1, issues=issues)

        result = apply_rollback([], validation, issues, orders)

        assert len(result.summary_entries) == 1
        entry = result.summary_entries[0]
        assert entry.status == STATUS_UNALLOCATED
        assert entry.reason == (
            "E01 - 关键字段为空或格式异常; "
            "E02 - 退单表行号重复或不连续; "
            "E03 - QC情况非法; "
            "E04 - 数量非法"
        )

    def test_e08_cascade_for_wms_without_sku(self):
        """E08 挂在 SKU1 上；不含 SKU1 的退单号 W2 只被连带回滚。"""
        orders = [
            make_order(excel_row_num=2, wms_order_no="W1", sku="SKU1"),
            make_order(excel_row_num=3, wms_order_no="W2", sku="SKU2"),
        ]
        issues = [
            make_issue(wms_order_no=NA_PLACEHOLDER, sku="SKU1", error_code=ERR_E08,
                       source_table=NA_PLACEHOLDER, excel_row_num=0),
        ]
        validation = ValidationResult(has_failures=True, failed_shipment_count=1, issues=issues)

        result = apply_rollback([], validation, issues, orders)

        reasons = {e.wms_order_no: e.reason for e in result.summary_entries}
        assert reasons["W1"] == "E08 - 同物流单号+SKU数量不一致"
        assert reasons["W2"] == "整单回滚（触发原因：E08）"


# -----------------------------------------------------------------------------
# apply_rollback：分配阶段失败 → 整单回滚（E09 / E10 / 连带回滚组）
# -----------------------------------------------------------------------------


class TestAllocationStageRollback:
    def test_e10_direct_and_cascade_reasons(self):
        """SKU1 组 E10 失败、SKU2 组连带回滚 → 整单回滚；
        W1（含 SKU1）为直接原因，W2（仅 SKU2）为连带回滚格式。"""
        shipment = make_shipment("SH1", [
            make_group("SH1", "SKU1", success=False, error_code=ERR_E10),
            make_group("SH1", "SKU2", success=False, error_code=ERROR_CASCADE_ROLLBACK),
        ])
        orders = [
            make_order(excel_row_num=2, wms_order_no="W1", sku="SKU1"),
            make_order(excel_row_num=3, wms_order_no="W2", sku="SKU2"),
        ]

        result = apply_rollback([shipment], EMPTY_VALIDATION, [], orders)

        assert result.details == []
        reasons = {e.wms_order_no: e.reason for e in result.summary_entries}
        assert all(e.status == STATUS_UNALLOCATED for e in result.summary_entries)
        assert reasons["W1"] == "E10 - 回溯超限"
        assert reasons["W2"] == "整单回滚（触发原因：E10）"

    def test_e09_shared_pool_both_wms_direct(self):
        """多退单号共享同一 SKU 库存池：该 SKU 组 E09 失败，
        两个含此 SKU 的退单号均为直接原因（共享库存池不可分割）。"""
        shipment = make_shipment("SH1", [
            make_group("SH1", "SKU1", success=False, error_code=ERR_E09),
        ])
        orders = [
            make_order(excel_row_num=2, wms_order_no="W1", sku="SKU1"),
            make_order(excel_row_num=3, wms_order_no="W2", sku="SKU1"),
        ]

        result = apply_rollback([shipment], EMPTY_VALIDATION, [], orders)

        reasons = {e.wms_order_no: e.reason for e in result.summary_entries}
        assert reasons["W1"] == "E09 - 分配路径穷尽"
        assert reasons["W2"] == "E09 - 分配路径穷尽"

    def test_failed_shipment_rolls_back_success_group_details(self):
        """同一物流单号下 SKU1 成功、SKU2 E09 失败 → 整单回滚，成功明细不输出。"""
        success_detail = make_detail(wms_order_no="W1", sku="SKU1")
        shipment = make_shipment("SH1", [
            make_group("SH1", "SKU1", success=True, details=[success_detail]),
            make_group("SH1", "SKU2", success=False, error_code=ERR_E09),
        ])
        orders = [
            make_order(excel_row_num=2, wms_order_no="W1", sku="SKU1"),
            make_order(excel_row_num=3, wms_order_no="W1", sku="SKU2"),
        ]

        result = apply_rollback([shipment], EMPTY_VALIDATION, [], orders)

        assert result.details == []
        assert len(result.summary_entries) == 1
        entry = result.summary_entries[0]
        assert entry.status == STATUS_UNALLOCATED
        # W1 同时含失败 SKU2 → 直接原因（成功的 SKU1 也随整单回滚撤回）
        assert entry.reason == "E09 - 分配路径穷尽"


# -----------------------------------------------------------------------------
# apply_rollback：成功路径与退单号状态聚合（TC-04 / TC-05）
# -----------------------------------------------------------------------------


class TestSuccessAggregation:
    def test_tc05_manual_infects_wms_status(self):
        """TC-05：行 00001 两种批号/效期组合 → 手工操作；行 00002 批量导入；
        退单号状态聚合 = 手工操作，且所有明细行的退单号状态字段相同。"""
        details = [
            make_detail(wms_order_no="TK00000051", line_no="00001", order_qty=2,
                        lot_no="LA03", expiry="2031/01/01", alloc_qty=1),
            make_detail(wms_order_no="TK00000051", line_no="00001", order_qty=2,
                        lot_no="LA02", expiry="2030/01/01", alloc_qty=1),
            make_detail(wms_order_no="TK00000051", line_no="00002", order_qty=1,
                        lot_no="LA01", expiry="2029/01/01", alloc_qty=1),
        ]
        shipment = make_shipment("SF3190000000005", [
            make_group("SF3190000000005", "H000000001", success=True, details=details),
        ])

        result = apply_rollback([shipment], EMPTY_VALIDATION, [], [])

        assert len(result.summary_entries) == 1
        entry = result.summary_entries[0]
        assert entry.wms_order_no == "TK00000051"
        assert entry.status == STATUS_MANUAL
        assert entry.reason == ""

        assert len(result.details) == 3
        by_line = {}
        for row in result.details:
            by_line.setdefault(row.line_no, []).append(row)
            assert row.wms_order_status == STATUS_MANUAL
        assert [r.alloc_qty for r in by_line["00001"]] == [1, 1]
        assert all(r.order_qty == 2 for r in by_line["00001"])
        assert all(r.line_status == STATUS_MANUAL for r in by_line["00001"])
        assert by_line["00002"][0].line_status == STATUS_BATCH_IMPORT

    def test_tc04_all_batch_multi_wms_shared_pool(self):
        """TC-04/TC-08：两个退单号共享同一 SKU 库存池且全部批量导入；
        两个退单号各自聚合为批量导入，互不污染。"""
        ship = "SF3190000000018"
        details = [
            make_detail(shipment_no=ship, wms_order_no="TK00000181", line_no="00002",
                        order_qty=8, qc="QC", alloc_qty=8),
            make_detail(shipment_no=ship, wms_order_no="TK00000181", line_no="00003",
                        order_qty=12, qc="ZP", alloc_qty=12),
            make_detail(shipment_no=ship, wms_order_no="TK00000181", line_no="00001",
                        order_qty=6, qc="QC", alloc_qty=6),
            make_detail(shipment_no=ship, wms_order_no="TK00000181", line_no="00004",
                        order_qty=5, qc="QC", alloc_qty=5),
            make_detail(shipment_no=ship, wms_order_no="TK00000184", line_no="00001",
                        order_qty=5, qc="NG", alloc_qty=5),
        ]
        shipment = make_shipment(ship, [
            make_group(ship, "H000000001", success=True, details=details),
        ])

        result = apply_rollback([shipment], EMPTY_VALIDATION, [], [])

        statuses = {e.wms_order_no: e.status for e in result.summary_entries}
        assert statuses == {
            "TK00000181": STATUS_BATCH_IMPORT,
            "TK00000184": STATUS_BATCH_IMPORT,
        }
        assert len(result.details) == 5
        assert all(r.line_status == STATUS_BATCH_IMPORT for r in result.details)
        assert all(r.wms_order_status == STATUS_BATCH_IMPORT for r in result.details)

    def test_same_line_no_in_different_wms_not_merged(self):
        """不同退单号都有行号 00001：W1 的 00001 两种组合（手工操作），
        W2 的 00001 一种组合（批量导入），行状态与退单号状态不得互相污染。"""
        details = [
            make_detail(wms_order_no="W1", line_no="00001", lot_no="LA01", alloc_qty=1),
            make_detail(wms_order_no="W1", line_no="00001", lot_no="LA02", alloc_qty=1),
            make_detail(wms_order_no="W2", line_no="00001", lot_no="LB01", alloc_qty=3),
        ]
        shipment = make_shipment("SH1", [
            make_group("SH1", "SKU1", success=True, details=details),
        ])

        result = apply_rollback([shipment], EMPTY_VALIDATION, [], [])

        statuses = {e.wms_order_no: e.status for e in result.summary_entries}
        assert statuses == {"W1": STATUS_MANUAL, "W2": STATUS_BATCH_IMPORT}
        w2_rows = [r for r in result.details if r.wms_order_no == "W2"]
        assert w2_rows[0].line_status == STATUS_BATCH_IMPORT
        assert w2_rows[0].wms_order_status == STATUS_BATCH_IMPORT


# -----------------------------------------------------------------------------
# apply_rollback：E07 孤立物流单号 → [N/A] 专用记录（§5.1）
# -----------------------------------------------------------------------------


class TestOrphanShipment:
    def test_e07_generates_na_record(self):
        issues = [
            make_issue(shipment_no="SH_ORPHAN", wms_order_no=NA_PLACEHOLDER,
                       sku="SKU1", error_code="E07",
                       source_table=SOURCE_INVENTORY_TABLE, excel_row_num=2),
        ]
        validation = ValidationResult(has_failures=True, failed_shipment_count=1, issues=issues)

        result = apply_rollback([], validation, issues, [])

        assert result.details == []
        assert len(result.summary_entries) == 1
        entry = result.summary_entries[0]
        assert entry.shipment_no == "SH_ORPHAN"
        assert entry.wms_order_no == NA_PLACEHOLDER
        assert entry.status == STATUS_UNALLOCATED
        assert entry.reason == "E07 - 物流单号仅存在于质检库存表"


# -----------------------------------------------------------------------------
# aggregate_wms_status
# -----------------------------------------------------------------------------


class TestAggregateWmsStatus:
    def test_none_returns_empty(self):
        assert aggregate_wms_status(None) == []

    def test_empty_result_returns_empty(self):
        assert aggregate_wms_status(FinalResult()) == []

    def test_returns_summary_entries(self):
        shipment = make_shipment("SH1", [
            make_group("SH1", "SKU1", success=True, details=[make_detail()]),
        ])
        result = apply_rollback([shipment], EMPTY_VALIDATION, [], [])

        entries = aggregate_wms_status(result)

        assert len(entries) == 1
        assert entries[0].shipment_no == "SH1"
        assert entries[0].wms_order_no == "W1"
        assert entries[0].status == STATUS_BATCH_IMPORT
        assert entries[0].reason == ""
