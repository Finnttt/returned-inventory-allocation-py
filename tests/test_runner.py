"""M15 运行编排端到端测试（对应 VBA modRunner.bas，需求 §5.6 / §6.7.3~§6.7.5）。

覆盖：Dry Run / Full Run 两个入口的五张输出表内容、E12 中止行为
（不生成输出表、仅追加运行历史）、E99 统一捕获后终止、build_run_stats 统计口径、
build_error_code_distribution 错误码分布格式。
"""

import pytest

from returned_inventory import runner
from returned_inventory.excel_input import InputError
from returned_inventory.excel_output import RUN_HISTORY_HEADERS
from returned_inventory.guards import E99Error
from returned_inventory.models import (
    DEBUG_LEVEL_SIMPLE,
    STATUS_BATCH_IMPORT,
    STATUS_UNALLOCATED,
    GroupAllocResult,
    GroupStats,
    NormalizedInventoryLine,
    NormalizedReturnLine,
    ShipmentAllocResult,
    ValidationIssue,
    ValidationResult,
)
from returned_inventory.runner import (
    E12_ABORT_FLAG,
    build_error_code_distribution,
    build_run_stats,
    run_full_allocation,
    run_validation_only,
)

from .wb_factory import build_test_workbook, simple_success_workbook

SHIP = "SF001"
WMS = "WMS001"


def _data_rows(ws):
    return [
        [cell.value for cell in row]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row)
    ]


def _history_rows(wb):
    return _data_rows(wb["运行历史记录表"])


# -----------------------------------------------------------------------------
# Dry Run（run_validation_only）
# -----------------------------------------------------------------------------


def test_dry_run_all_pass_produces_empty_summary():
    """通过校验的物流单号不出现在干跑汇总表中（需求 §6.7.5）。"""
    wb = simple_success_workbook()
    stats = run_validation_only(wb)

    assert stats.validation_fail_count == 0
    assert stats.input_return_rows == 1
    assert stats.input_inventory_rows == 1
    assert stats.input_shipment_count == 1
    assert stats.alloc_success_count == 0  # 干跑分配字段恒为 0

    assert wb["分配状态汇总表"].max_row == 1  # 仅表头
    assert wb["成功分配明细表"].max_row == 1
    assert wb["数据异常明细表"].max_row == 1
    assert wb["调试日志"].max_row == 1

    history = _history_rows(wb)
    assert len(history) == 1
    assert history[0][0] == 1           # 运行编号自增
    assert history[0][2] == "Dry Run"
    assert history[0][9] == 0           # 校验失败物流单号数
    assert history[0][12] == "" or history[0][12] is None  # 错误码分布留空


def test_dry_run_validation_failure_appears_in_summary():
    """E06：库存表缺该物流单号 → 汇总表含"无法分配"行，原因带错误码。"""
    wb = build_test_workbook(
        order_rows=[[SHIP, WMS, "SKU-A", "00001", 5]],
        inventory_rows=[],  # 无库存 → E06
    )
    stats = run_validation_only(wb)

    assert stats.validation_fail_count == 1
    summary = _data_rows(wb["分配状态汇总表"])
    assert len(summary) == 1
    assert summary[0][0] == SHIP
    assert summary[0][2] == STATUS_UNALLOCATED
    assert "E06" in summary[0][3]

    # E06 按路由变更逐行进入数据异常明细表
    anomaly = _data_rows(wb["数据异常明细表"])
    assert len(anomaly) == 1
    assert anomaly[0][7] == "E06"

    history = _history_rows(wb)
    assert history[0][12] == "E06:1"  # 错误码分布


# -----------------------------------------------------------------------------
# Full Run（run_full_allocation）
# -----------------------------------------------------------------------------


def test_full_run_success_writes_all_output_sheets():
    wb = simple_success_workbook()
    stats = run_full_allocation(wb)

    assert stats.alloc_success_count == 1
    assert stats.alloc_fail_count == 0
    assert stats.total_backtrack_count == 0

    summary = _data_rows(wb["分配状态汇总表"])
    assert summary == [[SHIP, WMS, STATUS_BATCH_IMPORT, None]] or summary == [[SHIP, WMS, STATUS_BATCH_IMPORT, ""]]

    detail = _data_rows(wb["成功分配明细表"])
    assert detail == [[
        SHIP, WMS, "SKU-A", "00001", 5, "ZP", "LA01",
        "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT,
    ]]

    assert wb["数据异常明细表"].max_row == 1
    assert wb["调试日志"].max_row == 1  # 默认配置调试日志关闭

    history = _history_rows(wb)
    assert len(history) == 1
    assert history[0][2] == "Full Run"
    assert history[0][10] == 1   # 分配成功物流单号数
    assert history[0][11] == 0   # 分配失败物流单号数
    assert history[0][13] == 0   # 总回溯次数
    assert history[0][18] == "不敏感"      # 批号比较模式快照
    assert history[0][19] == "2099/01/01"  # 无保质期哨兵值快照


def test_full_run_allocation_failure_and_debug_log_simple():
    """预检测B 场景（需求 {1,2,5,6}，库存 ZP=7/QC=4/NG=3）→ E09；
    简版调试日志每个退单行一条最终结果记录。"""
    wb = build_test_workbook(
        order_rows=[
            [SHIP, WMS, "SKU-A", "00001", 1],
            [SHIP, WMS, "SKU-A", "00002", 2],
            [SHIP, WMS, "SKU-A", "00003", 5],
            [SHIP, WMS, "SKU-A", "00004", 6],
        ],
        inventory_rows=[
            [SHIP, "SKU-A", "ZP", "LA01", "2029/01/01", 7],
            [SHIP, "SKU-A", "QC", "LB01", "2029/01/01", 4],
            [SHIP, "SKU-A", "NG", "LC01", "2029/01/01", 3],
        ],
        config_rows=[["调试日志级别", DEBUG_LEVEL_SIMPLE, ""]],
    )
    stats = run_full_allocation(wb)

    assert stats.validation_fail_count == 0  # E08 总数相等、E11 无碎片 → 通过前校验
    assert stats.alloc_fail_count == 1
    summary = _data_rows(wb["分配状态汇总表"])
    assert summary[0][2] == STATUS_UNALLOCATED
    assert "E09" in summary[0][3]
    assert wb["成功分配明细表"].max_row == 1  # 失败单不写成功明细

    debug = _data_rows(wb["调试日志"])
    assert len(debug) == 4                    # 每个退单行 1 条最终结果（简版）
    assert all(row[17] == "E09" for row in debug)  # 第 18 列错误码
    assert all(len(row) == 19 for row in debug)

    history = _history_rows(wb)
    assert history[0][12] == "E09:1"
    assert history[0][15] == DEBUG_LEVEL_SIMPLE


def test_full_run_rerun_overwrites_outputs_but_appends_history():
    """重跑覆盖（需求 §6.7.4）：输出表数据被覆盖，运行历史追加不清空。"""
    wb = simple_success_workbook()
    run_full_allocation(wb)
    run_full_allocation(wb)

    assert len(_history_rows(wb)) == 2
    assert _data_rows(wb["成功分配明细表"])  # 数据仍在（被覆盖而非累加）
    assert len(_data_rows(wb["成功分配明细表"])) == 1


# -----------------------------------------------------------------------------
# E12 中止行为
# -----------------------------------------------------------------------------


def test_e12_aborts_run_and_appends_history_only():
    """表头错位触发 E12：中止运行、不生成输出表、仅追加运行历史。"""
    wb = build_test_workbook(
        order_rows=[[SHIP, WMS, "SKU-A", "00001", 5]],
        inventory_rows=[[SHIP, "SKU-A", "ZP", "LA01", "2029/01/01", 5]],
        order_headers=["物流单号", "WMS退单号", "sku", "行号", "数量"],  # 第 3 列错位
    )

    with pytest.raises(InputError, match="表头校验失败"):
        run_validation_only(wb)

    # 不生成输出表数据
    for name in ("分配状态汇总表", "成功分配明细表", "数据异常明细表", "调试日志"):
        assert wb[name].max_row == 1

    history = _history_rows(wb)
    assert len(history) == 1
    assert history[0][0] == 1                 # 运行编号仍自增
    assert history[0][2] == "Dry Run"
    assert history[0][9] == E12_ABORT_FLAG    # 校验失败物流单号数填 [E12-中止]
    assert history[0][3] == 0                 # 其余数值字段填 0
    assert history[0][12] == "[N/A]"          # 错误码分布填 [N/A]


def test_e12_full_run_history_marks_full_run():
    wb = build_test_workbook(
        order_rows=[[SHIP, WMS, "SKU-A", "00001", 5]],
        inventory_rows=[[SHIP, "SKU-A", "ZP", "LA01", "2029/01/01", 5]],
        order_headers=["WMS退单号", "物流单号", "SKU", "行号", "数量"],
    )

    with pytest.raises(InputError):
        run_full_allocation(wb)

    history = _history_rows(wb)
    assert history[0][2] == "Full Run"
    assert history[0][9] == E12_ABORT_FLAG


# -----------------------------------------------------------------------------
# E99 统一捕获（对齐 VBA M15 E99Fail 分支）
# -----------------------------------------------------------------------------


def test_e99_aborts_without_output_and_history(monkeypatch):
    """守卫触发 E99：不生成输出表、不追加运行历史，异常向上抛。"""
    def _raise_e99(*args, **kwargs):
        raise E99Error(SHIP, "SKU-A", 5, 4, "AssertConservation")

    monkeypatch.setattr(runner, "allocate_shipment", _raise_e99)
    wb = simple_success_workbook()

    with pytest.raises(E99Error, match="E99"):
        run_full_allocation(wb)

    assert wb["分配状态汇总表"].max_row == 1
    assert len(_history_rows(wb)) == 0


# -----------------------------------------------------------------------------
# build_run_stats / build_error_code_distribution
# -----------------------------------------------------------------------------


def _group(success=True, error_code="", backtrack=0):
    return GroupAllocResult(
        shipment_no=SHIP,
        sku="SKU-A",
        success=success,
        error_code=error_code,
        stats=GroupStats(SHIP, "SKU-A", backtrack, ""),
    )


def _order_line(ship=SHIP):
    return NormalizedReturnLine(
        excel_row_num=2,
        shipment_no=ship,
        wms_order_no=WMS,
        sku="SKU-A",
        line_no="00001",
        qty=5,
        line_no_valid=True,
        qty_valid=True,
        empty_fields="",
    )


def _inventory_line(ship=SHIP):
    return NormalizedInventoryLine(
        excel_row_num=2,
        shipment_no=ship,
        sku="SKU-A",
        qc="ZP",
        lot_no="LA01",
        expiry="2029/01/01",
        qty=5,
        qc_valid=True,
        expiry_valid=True,
        qty_valid=True,
        empty_fields="",
    )


def test_build_run_stats_aggregates_groups():
    results = [
        ShipmentAllocResult("SF001", [_group(backtrack=3), _group(backtrack=1)]),
        ShipmentAllocResult("SF002", [_group(success=False, error_code="E09")]),
        ShipmentAllocResult("SF003", []),  # 无 SKU 组：不计入成功/失败
    ]
    validation = ValidationResult(has_failures=True, failed_shipment_count=2)
    stats = build_run_stats(
        validation,
        results,
        orders=[_order_line(), _order_line(), _order_line()],
        inventory=[_inventory_line()],
    )

    assert stats.input_return_rows == 3
    assert stats.input_inventory_rows == 1
    assert stats.input_shipment_count == 1
    assert stats.validation_fail_count == 2
    assert stats.alloc_success_count == 1
    assert stats.alloc_fail_count == 1
    assert stats.total_backtrack_count == 4
    assert stats.max_group_backtrack == 3


def test_build_run_stats_dry_run_empty_results():
    validation = ValidationResult(has_failures=False, failed_shipment_count=0)
    stats = build_run_stats(validation, [], [], [])
    assert stats.alloc_success_count == 0
    assert stats.alloc_fail_count == 0
    assert stats.total_backtrack_count == 0
    assert stats.input_shipment_count == 0


def _issue(code, ship):
    return ValidationIssue(
        shipment_no=ship,
        wms_order_no=WMS,
        sku="SKU-A",
        error_code=code,
        source_table="退单表",
        excel_row_num=2,
        field_name="",
        raw_value="",
        reason="",
    )


def test_error_code_distribution_format_and_dedup():
    """格式 `E01:3; E09:1`：按码升序、错误码+物流单号去重计数。"""
    issues = [
        _issue("E01", "SF001"),
        _issue("E01", "SF001"),  # 同码同单号 → 去重
        _issue("E01", "SF002"),
        _issue("E01", "SF003"),
        _issue("E04", "SF002"),
    ]
    results = [ShipmentAllocResult("SF009", [_group(success=False, error_code="E09")])]

    assert build_error_code_distribution(issues, results) == "E01:3; E04:1; E09:1"


def test_error_code_distribution_empty_when_no_errors():
    assert build_error_code_distribution([], []) == ""
