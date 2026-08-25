"""M15 运行编排（对应 VBA modRunner.bas，需求 §5.6 / §6.7.3~§6.7.5）。

职责：系统的两个入口（干跑 / 完整运行），串联所有模块的调用顺序；
以及 build_run_stats / build_error_code_distribution，统一构造运行统计。

- run_validation_only（Dry Run）：清空 → 加载 → 标准化 → 前校验 →
  build_run_stats → 构建输出 → 写入。只输出校验失败物流单号的汇总行，
  不产生成功明细与调试日志（需求 §6.7.5）。
- run_full_allocation（Full Run）：…建账本 → 排序预检 → 回溯分配 →
  整单状态 → 分配后校验（assert_post_valid）→ 输出。
- build_run_stats：干跑时 shipment_results 传空列表，分配相关字段均为 0。

异常处理（与 VBA 对齐）：
- ConfigError：配置读取失败，原样向上抛，由 CLI 提示"请修正 输入_配置 后重试"。
- InputError（E12 结构异常）：数据加载阶段中止整次运行，不生成任何输出表；
  按需求 §4.1 第 0 层向运行历史记录表追加一条记录（校验失败物流单号数填
  [E12-中止]，其余字段填 0 或 [N/A]），随后原样向上抛。
  注意：VBA modRunner 实际并未实现 E12 历史追加（RunFail 仅弹窗），此处按
  需求文档 §4.1 / §5.6 实现，属于对 VBA 的有意补齐。
- E99Error：库存守恒等式被破坏的严重错误。对应 VBA 的 E99Fail 分支——
  不生成输出表、不追加运行历史，原样向上抛，由 CLI 报告后终止。

与 VBA 的结构差异：
- VBA 的 ShipmentAllocResult 为 Scripting.Dictionary（GroupCount / Group_g_* 平铺键）；
  Python 用 ShipmentAllocResult / GroupAllocResult dataclass，字段一一对应。
- VBA 用 Timer 计时（处理午夜回绕）；Python 用 time.perf_counter，
  耗时同样保留 1 位小数（RN_ElapsedSecs 口径）。
"""

from __future__ import annotations

import time
from datetime import datetime
from typing import Any

from .backtracking import allocate_shipment
from .config import load_config
from .excel_input import InputError, read_qc_inventory, read_return_orders
from .excel_output import (
    ANOMALY_HEADERS,
    DETAIL_HEADERS,
    SUMMARY_HEADERS,
    append_run_history,
    clear_output_sheets,
    write_debug_log,
    write_sheet,
)
from .ledger import InventoryLedger, build_ledger
from .models import (
    NA_PLACEHOLDER,
    AllocationEvent,
    Config,
    NormalizedInventoryLine,
    NormalizedReturnLine,
    PrecheckResult,
    RunStats,
    ShipmentAllocResult,
    ValidationIssue,
    ValidationResult,
)
from .normalize import normalize_inventory_rows, normalize_return_rows
from .output_builder import (
    build_anomaly_output_rows,
    build_debug_log_rows,
    build_detail_rows,
    build_run_history_row,
    build_summary_rows,
)
from .post_validate import assert_post_valid, validate_post
from .sort_filter import build_static_plan, run_precheck
from .status import FinalResult, aggregate_wms_status, apply_rollback
from .validate import build_anomaly_rows, validate_pre

# 输入表 / 配置表 / 运行历史表名称（与 VBA modRunner 私有常量一致）
SHEET_RETURN_INPUT = "输入_退单表"
SHEET_INVENTORY_INPUT = "输入_质检库存表"
SHEET_CONFIG = "输入_配置"
SHEET_RUN_HISTORY = "运行历史记录表"

# E12 中止时运行历史"校验失败物流单号数"列的填值（需求 §4.1 第 0 层）
E12_ABORT_FLAG = "[E12-中止]"


# -----------------------------------------------------------------------------
# 一、公开函数：两个入口
# -----------------------------------------------------------------------------


def run_validation_only(wb: Any) -> RunStats:
    """干跑入口：只做校验，不执行分配（VBA RunValidationOnly，需求 §6.7.5）。

    流程：清空输出 → 读取原始数据 → 标准化 → 分配前校验 →
    build_run_stats（分配字段为 0）→ 构建输出行 → 写入 Excel。
    返回本次运行统计；ConfigError / InputError / OutputError 向上抛。
    """
    if wb is None:
        raise ValueError("工作簿对象为空，无法运行。")

    run_start_text = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    t0 = time.perf_counter()

    cfg = load_config(_get_sheet_or_none(wb, SHEET_CONFIG))

    clear_output_sheets(wb, cfg)

    try:
        raw_orders = read_return_orders(_get_sheet_or_none(wb, SHEET_RETURN_INPUT))
        raw_inventory = read_qc_inventory(_get_sheet_or_none(wb, SHEET_INVENTORY_INPUT))
    except InputError:
        _append_e12_run_history(wb, run_start_text, dry_run_mode=True)
        raise

    orders, return_issues = normalize_return_rows(raw_orders, cfg)
    inventory, inventory_issues = normalize_inventory_rows(raw_inventory, cfg)
    normalized_issues = return_issues + inventory_issues

    validation_result = validate_pre(orders, inventory, normalized_issues, cfg)

    # 干跑：分配结果传空列表，分配相关字段全为 0
    stats = build_run_stats(validation_result, [], orders, inventory)

    # 干跑的 FinalResult 只有校验失败项，无成功分配明细
    final_result = apply_rollback([], validation_result, validation_result.issues, orders)

    elapsed = _elapsed_secs(t0, time.perf_counter())
    run_history_row = build_run_history_row(
        stats, cfg, True, run_start_text, elapsed, 0, elapsed,
        build_error_code_distribution(validation_result.issues, []),
    )

    _write_all_output(wb, cfg, final_result, validation_result.issues, [], run_history_row, True)
    return stats


def run_full_allocation(wb: Any) -> RunStats:
    """完整运行入口：在通过校验的物流单号上执行回溯分配（VBA RunFullAllocation）。

    流程：清空 → 读取 → 标准化 → 前校验 → 建账本 → 排序预检 → 回溯分配 →
    整单状态判定 → 分配后校验（失败升级 E99）→ build_run_stats → 输出 → 写入。
    E99Error 不在此吞掉：对应 VBA E99Fail 分支，不生成输出表，向上抛由 CLI 报告。
    """
    if wb is None:
        raise ValueError("工作簿对象为空，无法运行。")

    run_start_text = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    t0 = time.perf_counter()

    cfg = load_config(_get_sheet_or_none(wb, SHEET_CONFIG))

    clear_output_sheets(wb, cfg)

    try:
        raw_orders = read_return_orders(_get_sheet_or_none(wb, SHEET_RETURN_INPUT))
        raw_inventory = read_qc_inventory(_get_sheet_or_none(wb, SHEET_INVENTORY_INPUT))
    except InputError:
        _append_e12_run_history(wb, run_start_text, dry_run_mode=False)
        raise

    orders, return_issues = normalize_return_rows(raw_orders, cfg)
    inventory, inventory_issues = normalize_inventory_rows(raw_inventory, cfg)
    normalized_issues = return_issues + inventory_issues

    validation_result = validate_pre(orders, inventory, normalized_issues, cfg)
    t_validate = time.perf_counter()

    # 建账本（M06）；分配循环中守卫断言失败抛 E99Error，沿调用链直传 CLI
    ledger = build_ledger(inventory)

    # 回溯分配（M07 + M09）：仅对通过校验的物流单号执行
    shipment_results = _run_all_allocations(orders, ledger, cfg, validation_result.issues)
    t_alloc = time.perf_counter()

    # 整单状态判定（M11）
    final_result = apply_rollback(shipment_results, validation_result, validation_result.issues, orders)

    # 分配后校验（M12）：写入成功明细前的最后一道防线，失败即抛 E99Error
    assert_post_valid(
        validate_post(orders, final_result.details, final_result.summary_entries)
    )

    stats = build_run_stats(validation_result, shipment_results, orders, inventory)

    # 从分配结果中合并调试日志事件（M09 → M15 → M13）
    events = _collect_debug_events(shipment_results)

    run_history_row = build_run_history_row(
        stats, cfg, False, run_start_text,
        _elapsed_secs(t0, t_validate),
        _elapsed_secs(t_validate, t_alloc),
        _elapsed_secs(t0, time.perf_counter()),
        build_error_code_distribution(validation_result.issues, shipment_results),
    )

    _write_all_output(wb, cfg, final_result, validation_result.issues, events, run_history_row, False)
    return stats


# -----------------------------------------------------------------------------
# 二、公开函数：build_run_stats / build_error_code_distribution
# -----------------------------------------------------------------------------


def build_run_stats(
    validation_result: ValidationResult,
    shipment_results: list[ShipmentAllocResult],
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
) -> RunStats:
    """统一构造本次运行的汇总统计（VBA BuildRunStats），供运行历史记录表使用。

    干跑时 shipment_results 传空列表：total_backtrack_count / max_group_backtrack /
    alloc_success_count / alloc_fail_count 均为 0。
    """
    stats = RunStats(
        input_return_rows=len(orders),
        input_inventory_rows=len(inventory),
        input_shipment_count=_count_distinct_shipments(orders, inventory),
        validation_fail_count=validation_result.failed_shipment_count,
    )

    for result in shipment_results:
        # 跳过无 SKU 组的条目（对应 VBA GroupCount=0 防御分支）
        if result is None or len(result.group_results) == 0:
            continue

        ship_all_success = True
        for group in result.group_results:
            backtrack = group.stats.backtrack_count
            stats.total_backtrack_count += backtrack
            if backtrack > stats.max_group_backtrack:
                stats.max_group_backtrack = backtrack
            if not group.success:
                ship_all_success = False

        # 整单成功/失败以物流单号为粒度：所有 SKU 组均成功才算整单成功
        if ship_all_success:
            stats.alloc_success_count += 1
        else:
            stats.alloc_fail_count += 1

    return stats


def build_error_code_distribution(
    validation_issues: list[ValidationIssue],
    shipment_results: list[ShipmentAllocResult],
) -> str:
    """汇总各错误码命中的物流单号数（VBA RN_BuildErrorCodeDistribution）。

    格式 "E01:3; E09:1"：先按 错误码|物流单号 去重，再按错误码字符串升序计数；
    无错误时返回空串。分配阶段的连带回滚组（error_code="连带回滚"）在 VBA 中
    同样计入分布（Group_g_ErrorCode 非空即计数），此处原样保留该口径。
    """
    pair_seen: set[str] = set()
    code_counts: dict[str, int] = {}

    for issue in validation_issues:
        code = issue.error_code.strip()
        ship = issue.shipment_no.strip()
        if code and ship:
            pair_key = f"{code}|{ship}"
            if pair_key not in pair_seen:
                pair_seen.add(pair_key)
                code_counts[code] = code_counts.get(code, 0) + 1

    for result in shipment_results:
        if result is None:
            continue
        ship_no = result.shipment_no
        for group in result.group_results:
            code = group.error_code.strip()
            if code and ship_no:
                pair_key = f"{code}|{ship_no}"
                if pair_key not in pair_seen:
                    pair_seen.add(pair_key)
                    code_counts[code] = code_counts.get(code, 0) + 1

    if not code_counts:
        return ""
    return "; ".join(f"{code}:{code_counts[code]}" for code in sorted(code_counts))


# -----------------------------------------------------------------------------
# 三、私有函数：分配编排
# -----------------------------------------------------------------------------


def _run_all_allocations(
    orders: list[NormalizedReturnLine],
    ledger: InventoryLedger,
    cfg: Config,
    validation_issues: list[ValidationIssue],
) -> list[ShipmentAllocResult]:
    """对所有通过校验的物流单号执行回溯分配（VBA RN_RunAllAllocations）。

    按"物流单号 → SKU"两层分组，依次为每个 SKU 调用 build_static_plan /
    run_precheck（M07），再调用 allocate_shipment（M09，内部处理短路和连带回滚）。
    物流单号/SKU 的去重按大小写不敏感口径（对应 VBA Dictionary vbTextCompare），
    保留首次出现的原始写法。
    """
    ship_nos = _collect_unique([order.shipment_no for order in orders])

    # 校验阶段已失败的物流单号不参与分配（与 M11 apply_rollback 口径一致）
    failed_shipments = {
        issue.shipment_no.lower() for issue in validation_issues if issue.shipment_no
    }
    ship_nos = [sno for sno in ship_nos if sno.lower() not in failed_shipments]

    results: list[ShipmentAllocResult] = []
    for ship_no in ship_nos:
        sku_list, plan_map, precheck_map = _build_sku_groups_for_shipment(
            orders, ship_no, ledger
        )
        results.append(
            allocate_shipment(ship_no, sku_list, plan_map, precheck_map, ledger, cfg)
        )
    return results


def _build_sku_groups_for_shipment(
    orders: list[NormalizedReturnLine],
    ship_no: str,
    ledger: InventoryLedger,
) -> tuple[list[str], dict, dict]:
    """为指定物流单号构建 skuList / planMap / precheckMap（VBA RN_BuildSkuGroupsForShipment）。"""
    sku_list = _collect_unique(
        [order.sku for order in orders if order.shipment_no == ship_no]
    )

    plan_map: dict[str, Any] = {}
    precheck_map: dict[str, PrecheckResult] = {}
    for sku in sku_list:
        group_rows = [
            order
            for order in orders
            if order.shipment_no == ship_no and order.sku == sku
        ]
        plan = build_static_plan(group_rows, ledger)
        plan_map[sku] = plan
        precheck_map[sku] = run_precheck(plan, ledger)
    return sku_list, plan_map, precheck_map


def _collect_unique(values: list[str]) -> list[str]:
    """大小写不敏感去重、保插入序（对应 VBA Dictionary + vbTextCompare 的收集方式）。"""
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        key = value.lower()
        if key not in seen:
            seen.add(key)
            result.append(value)
    return result


def _collect_debug_events(
    shipment_results: list[ShipmentAllocResult],
) -> list[AllocationEvent]:
    """从所有物流单号分配结果中合并调试日志事件（VBA RN_CollectDebugEvents）。"""
    events: list[AllocationEvent] = []
    for result in shipment_results:
        if result is None:
            continue
        for group in result.group_results:
            events.extend(group.events)
    return events


# -----------------------------------------------------------------------------
# 四、私有函数：输出写入与运行历史
# -----------------------------------------------------------------------------


def _write_all_output(
    wb: Any,
    cfg: Config,
    final_result: FinalResult,
    validation_issues: list[ValidationIssue],
    events: list[AllocationEvent],
    run_history_row: list,
    dry_run_mode: bool,
) -> None:
    """统一写入所有输出工作表（VBA RN_WriteAllOutput）。

    干跑模式下汇总表只含校验失败项；明细与调试日志本就没有数据行
    （final_result 无 details、events 为空），仍写入表头保持表结构完整。
    """
    status_map = aggregate_wms_status(final_result)
    anomaly_rows = build_anomaly_rows(validation_issues)

    write_sheet(
        wb["分配状态汇总表"],
        build_summary_rows(status_map, dry_run_mode),
        SUMMARY_HEADERS,
    )
    write_sheet(
        wb["成功分配明细表"],
        build_detail_rows(final_result),
        DETAIL_HEADERS,
    )
    write_sheet(
        wb["数据异常明细表"],
        build_anomaly_output_rows(anomaly_rows),
        ANOMALY_HEADERS,
    )
    write_debug_log(wb, build_debug_log_rows(events, cfg), cfg)

    # 运行历史每次追加一行，不覆盖历史记录
    append_run_history(wb[SHEET_RUN_HISTORY], run_history_row)


def _append_e12_run_history(wb: Any, run_start_text: str, dry_run_mode: bool) -> None:
    """E12 中止时追加的运行历史记录（需求 §4.1 第 0 层 / §5.6）。

    校验失败物流单号数填 [E12-中止]，其余数值字段填 0、文本字段填 [N/A]、
    备注留空；运行编号仍由 append_run_history 自增生成。
    若运行历史表缺失或受保护，让 OutputError 向上抛（E12 已由调用方重新抛出，
    历史追加失败属于更严重的结构问题）。
    """
    row = [
        "",
        run_start_text,
        "Dry Run" if dry_run_mode else "Full Run",
        0,
        0,
        0,
        0,
        0,
        0,
        E12_ABORT_FLAG,
        0,
        0,
        NA_PLACEHOLDER,
        0,
        0,
        NA_PLACEHOLDER,
        "",
        NA_PLACEHOLDER,
        NA_PLACEHOLDER,
        NA_PLACEHOLDER,
    ]
    append_run_history(wb[SHEET_RUN_HISTORY], row)


# -----------------------------------------------------------------------------
# 五、私有工具函数
# -----------------------------------------------------------------------------


def _get_sheet_or_none(wb: Any, sheet_name: str) -> Any:
    """按名取工作表，缺失时返回 None（交由下游模块抛带中文文案的错误）。"""
    return wb[sheet_name] if sheet_name in wb.sheetnames else None


def _count_distinct_shipments(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
) -> int:
    """两表去重合并后的物流单号总数（VBA RN_CountDistinctShipments，需求 §5.6）。"""
    seen: set[str] = set()
    for line in orders:
        if line.shipment_no:
            seen.add(line.shipment_no.lower())
    for line in inventory:
        if line.shipment_no:
            seen.add(line.shipment_no.lower())
    return len(seen)


def _elapsed_secs(t_start: float, t_end: float) -> float:
    """耗时秒数，保留 1 位小数（VBA RN_ElapsedSecs 口径；perf_counter 无午夜回绕）。"""
    return round(t_end - t_start, 1)
