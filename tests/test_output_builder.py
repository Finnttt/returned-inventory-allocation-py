"""M13 输出构建单元测试（对应 VBA modOutputBuilder.bas，需求 §5.1~§5.6）。

覆盖：五张表的构建逻辑——汇总表干跑差异与 [N/A] 占位、明细表 11 列、
异常明细的错误码路由（E06 按 2026-07-20 路由变更进入本表；E08/E11 仅进汇总）、
调试日志三档级别过滤与 19 列结构、运行历史 20 列单行、空输入不崩溃。
"""

from returned_inventory.models import (
    DEBUG_LEVEL_DETAIL,
    DEBUG_LEVEL_OFF,
    DEBUG_LEVEL_SIMPLE,
    NA_PLACEHOLDER,
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STATUS_UNALLOCATED,
    AllocationEvent,
    AnomalyRow,
    Config,
    RunStats,
    WMSStatusEntry,
)
from returned_inventory.output_builder import (
    build_anomaly_output_rows,
    build_debug_log_rows,
    build_detail_rows,
    build_run_history_row,
    build_summary_rows,
)
from returned_inventory.status import FinalDetailRow, FinalResult

# -----------------------------------------------------------------------------
# 辅助构造
# -----------------------------------------------------------------------------


def _status_entry(ship="SF001", wms="WMS001", status=STATUS_BATCH_IMPORT, reason=""):
    return WMSStatusEntry(shipment_no=ship, wms_order_no=wms, status=status, reason=reason)


def _detail_row(**overrides):
    values = dict(
        shipment_no="SF001",
        wms_order_no="WMS001",
        sku="SKU-A",
        line_no="00001",
        order_qty=5,
        qc="ZP",
        lot_no="LA01",
        expiry="2029/01/01",
        alloc_qty=5,
        line_status=STATUS_BATCH_IMPORT,
        wms_order_status=STATUS_BATCH_IMPORT,
    )
    values.update(overrides)
    return FinalDetailRow(**values)


def _anomaly(error_code="E01", **overrides):
    values = dict(
        source_table="退单表",
        excel_row_num=2,
        shipment_no="SF001",
        wms_order_no="WMS001",
        sku="SKU-A",
        field_name="数量",
        raw_value="abc",
        error_code=error_code,
        reason="数量非法",
    )
    values.update(overrides)
    return AnomalyRow(**values)


def _event(is_final=True, **overrides):
    values = dict(
        shipment_no="SF001",
        sku="SKU-A",
        wms_order_no="WMS001",
        line_no="00001",
        demand_d=5,
        process_order="1",
        dynamic_next_min_qty="",
        candidate_qc_count="1",
        excluded_qc_list="",
        strategy_used="策略一",
        used_qc="ZP",
        qc_before="-",
        qc_after="-",
        lot_expiry_combo_count="1",
        is_backtrack_retry="否",
        backtrack_no=0,
        line_status=STATUS_BATCH_IMPORT,
        error_code="",
        fail_sub_type="",
        is_final_result=is_final,
        is_revoked=False,
    )
    values.update(overrides)
    return AllocationEvent(**values)


# -----------------------------------------------------------------------------
# 分配状态汇总表（§5.1）
# -----------------------------------------------------------------------------


def test_summary_rows_full_run_includes_success_and_failure():
    entries = [
        _status_entry(status=STATUS_BATCH_IMPORT),
        _status_entry(ship="SF002", status=STATUS_UNALLOCATED, reason="E09 - 分配路径穷尽"),
    ]
    rows = build_summary_rows(entries, dry_run_mode=False)
    assert rows == [
        ["SF001", "WMS001", STATUS_BATCH_IMPORT, ""],
        ["SF002", "WMS001", STATUS_UNALLOCATED, "E09 - 分配路径穷尽"],
    ]


def test_summary_rows_dry_run_keeps_only_unallocated():
    entries = [
        _status_entry(status=STATUS_BATCH_IMPORT),
        _status_entry(ship="SF002", status=STATUS_MANUAL),
        _status_entry(ship="SF003", status=STATUS_UNALLOCATED, reason="E06 - 物流单号仅存在于退单表"),
    ]
    rows = build_summary_rows(entries, dry_run_mode=True)
    assert rows == [["SF003", "WMS001", STATUS_UNALLOCATED, "E06 - 物流单号仅存在于退单表"]]


def test_summary_rows_empty_wms_falls_back_to_na_placeholder():
    """E07 孤立物流单号无 WMS 退单号，汇总表填 [N/A]。"""
    rows = build_summary_rows(
        [_status_entry(wms="", status=STATUS_UNALLOCATED, reason="E07 - 物流单号仅存在于质检库存表")],
        dry_run_mode=True,
    )
    assert rows[0][1] == NA_PLACEHOLDER


def test_summary_rows_empty_input_returns_empty():
    assert build_summary_rows([], dry_run_mode=False) == []
    assert build_summary_rows([], dry_run_mode=True) == []


# -----------------------------------------------------------------------------
# 成功分配明细表（§5.2）
# -----------------------------------------------------------------------------


def test_detail_rows_column_order():
    final = FinalResult(details=[_detail_row()])
    rows = build_detail_rows(final)
    assert rows == [[
        "SF001", "WMS001", "SKU-A", "00001", 5, "ZP", "LA01",
        "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT,
    ]]


def test_detail_rows_none_or_empty_returns_empty():
    assert build_detail_rows(None) == []
    assert build_detail_rows(FinalResult()) == []


# -----------------------------------------------------------------------------
# 数据异常明细表（§5.4）
# -----------------------------------------------------------------------------


def test_anomaly_rows_column_order():
    rows = build_anomaly_output_rows([_anomaly()])
    assert rows == [[
        "退单表", 2, "SF001", "WMS001", "SKU-A", "数量", "abc", "E01", "数量非法",
    ]]


def test_anomaly_rows_e08_e11_excluded_e06_kept():
    """E08/E11 只进汇总；E06 按路由变更与 E07 对称逐行进入异常明细。"""
    anomalies = [
        _anomaly(error_code="E06", reason="物流单号仅存在于退单表"),
        _anomaly(error_code="E07", source_table="质检库存表", reason="物流单号仅存在于质检库存表"),
        _anomaly(error_code="E08"),
        _anomaly(error_code="E11"),
    ]
    rows = build_anomaly_output_rows(anomalies)
    assert [row[7] for row in rows] == ["E06", "E07"]


def test_anomaly_rows_empty_input_returns_empty():
    assert build_anomaly_output_rows([]) == []


# -----------------------------------------------------------------------------
# 调试日志表（19 列，调试日志19列规格说明.md）
# -----------------------------------------------------------------------------


def test_debug_log_off_level_produces_no_rows():
    cfg = Config(debug_log_level=DEBUG_LEVEL_OFF)
    assert build_debug_log_rows([_event()], cfg) == []


def test_debug_log_simple_level_keeps_only_final_results():
    cfg = Config(debug_log_level=DEBUG_LEVEL_SIMPLE)
    events = [_event(is_final=False, line_status="过程-尝试成功"), _event(is_final=True)]
    rows = build_debug_log_rows(events, cfg)
    assert len(rows) == 1
    assert rows[0][16] == STATUS_BATCH_IMPORT  # 第 17 列：行状态


def test_debug_log_detail_level_keeps_all_events():
    cfg = Config(debug_log_level=DEBUG_LEVEL_DETAIL)
    events = [_event(is_final=False, line_status="过程-回溯撤销", is_revoked=True), _event()]
    rows = build_debug_log_rows(events, cfg)
    assert len(rows) == 2
    assert rows[0][16] == "过程-回溯撤销"


def test_debug_log_row_has_19_columns_in_spec_order():
    cfg = Config(debug_log_level=DEBUG_LEVEL_SIMPLE)
    rows = build_debug_log_rows([_event()], cfg)
    assert rows == [[
        "SF001", "SKU-A", "WMS001", "00001", 5, "1", "", "1", "",
        "策略一", "ZP", "-", "-", "1", "否", 0, STATUS_BATCH_IMPORT, "", "",
    ]]


def test_debug_log_empty_events_do_not_crash():
    for level in (DEBUG_LEVEL_OFF, DEBUG_LEVEL_SIMPLE, DEBUG_LEVEL_DETAIL):
        assert build_debug_log_rows([], Config(debug_log_level=level)) == []


# -----------------------------------------------------------------------------
# 运行历史记录表（§5.6，20 列）
# -----------------------------------------------------------------------------


def test_run_history_row_20_columns_dry_run():
    stats = RunStats(
        input_return_rows=3,
        input_inventory_rows=4,
        input_shipment_count=2,
        validation_fail_count=1,
    )
    cfg = Config()
    row = build_run_history_row(
        stats, cfg, True, "2026/08/22 10:00:00", 0.3, 0, 0.3, "E06:1"
    )
    assert row == [
        "",                      # 运行编号（由 append_run_history 自增填充）
        "2026/08/22 10:00:00",   # 运行时间
        "Dry Run",               # 运行类型
        3, 4, 2,                 # 输入行数/物流单号数
        0.3, 0, 0.3,             # 校验/分配/总耗时
        1, 0, 0,                 # 校验失败/分配成功/分配失败
        "E06:1",                 # 错误码分布
        0, 0,                    # 总回溯次数/最大单组回溯次数
        "关闭",                  # 调试日志级别
        "",                      # 备注
        200,                     # 最大回溯次数（配置快照）
        "不敏感",                # 批号比较模式（配置快照）
        "2099/01/01",            # 无保质期哨兵值（配置快照）
    ]


def test_run_history_row_full_run_config_snapshot():
    stats = RunStats(alloc_success_count=2, alloc_fail_count=1, total_backtrack_count=7, max_group_backtrack=5)
    cfg = Config(debug_log_level=DEBUG_LEVEL_SIMPLE, max_backtrack_count=50, lot_case_sensitive=True, no_expiry_sentinel="9999/12/31")
    row = build_run_history_row(stats, cfg, False, "2026/08/22 10:00:00", 0.1, 0.2, 0.3, "E09:1")
    assert row[2] == "Full Run"
    assert row[10:15] == [2, 1, "E09:1", 7, 5]
    assert row[15] == DEBUG_LEVEL_SIMPLE
    assert row[17:] == [50, "敏感", "9999/12/31"]
