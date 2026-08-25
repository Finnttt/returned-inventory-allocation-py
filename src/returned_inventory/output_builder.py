"""M13 输出构建（对应 VBA modOutputBuilder.bas，需求 §5.1~§5.6）。

职责：把领域层结果（WMSStatusEntry / FinalResult / AnomalyRow / AllocationEvent /
RunStats）翻译成"可直接写表"的输出行（OutputRow = list），本模块不做任何 Excel 写入。

与 VBA 的结构差异：
- VBA 的 BuildDetailRows 接收 FinalResult Scripting.Dictionary（DetailCount /
  Detail_i_* 平铺键）；Python 直接接收 status.FinalResult dataclass，
  details 列表与 Detail_i_* 一一对应。
- VBA 空数组以"未初始化数组 + UBound 报错"表达；Python 统一用空列表。
- 输出文案（[N/A] 占位符、Dry Run/Full Run、批号比较模式文本）与 VBA 逐字一致。
"""

from __future__ import annotations

from .models import (
    DEBUG_LEVEL_DETAIL,
    DEBUG_LEVEL_OFF,
    DEBUG_LEVEL_SIMPLE,
    ERR_E08,
    ERR_E11,
    LOT_MODE_INSENSITIVE,
    LOT_MODE_SENSITIVE,
    NA_PLACEHOLDER,
    STATUS_UNALLOCATED,
    AllocationEvent,
    AnomalyRow,
    Config,
    OutputRow,
    RunStats,
    WMSStatusEntry,
)
from .status import FinalResult


# -----------------------------------------------------------------------------
# 公开函数（与 VBA modOutputBuilder 公开函数一一对应）
# -----------------------------------------------------------------------------


def build_summary_rows(
    wms_status_map: list[WMSStatusEntry], dry_run_mode: bool
) -> list[OutputRow]:
    """构建"分配状态汇总表"行数组（VBA BuildSummaryRows，需求 §5.1）。

    dry_run_mode=True 时只保留校验失败（无法分配）项：通过校验但尚未分配的
    物流单号不进汇总，避免被误读为"已分配"（需求 §6.7.5）。
    """
    rows: list[OutputRow] = []
    for entry in wms_status_map:
        if dry_run_mode and entry.status != STATUS_UNALLOCATED:
            continue
        rows.append(
            [
                entry.shipment_no,
                _normalize_summary_wms_no(entry.wms_order_no),
                entry.status,
                entry.reason,
            ]
        )
    return rows


def build_detail_rows(final_result: FinalResult | None) -> list[OutputRow]:
    """构建"成功分配明细表"行数组（VBA BuildDetailRows，需求 §5.2，11 列）。

    仅整单成功的物流单号有明细；被整单回滚的物流单号由 status.apply_rollback
    拦截，不会出现在 final_result.details 中。
    """
    if final_result is None:
        return []
    return [
        [
            detail.shipment_no,
            detail.wms_order_no,
            detail.sku,
            detail.line_no,
            detail.order_qty,
            detail.qc,
            detail.lot_no,
            detail.expiry,
            detail.alloc_qty,
            detail.line_status,
            detail.wms_order_status,
        ]
        for detail in final_result.details
    ]


def build_anomaly_output_rows(anomaly_rows: list[AnomalyRow]) -> list[OutputRow]:
    """构建"数据异常明细表"行数组（VBA BuildAnomalyOutputRows，需求 §5.4，9 列）。

    E08/E11 只进汇总、不进入异常明细（跨行汇总比较，无具体异常字段）。
    E06 按 2026-07-20 路由变更与 E07 对称逐行进入本表（与 VBA
    OB_IsSummaryOnlyError 口径一致：仅 E08/E11 被排除）。
    """
    rows: list[OutputRow] = []
    for anomaly in anomaly_rows:
        if _is_summary_only_error(anomaly.error_code):
            continue
        rows.append(
            [
                anomaly.source_table,
                anomaly.excel_row_num,
                anomaly.shipment_no,
                anomaly.wms_order_no,
                anomaly.sku,
                anomaly.field_name,
                anomaly.raw_value,
                anomaly.error_code,
                anomaly.reason,
            ]
        )
    return rows


def build_debug_log_rows(
    events: list[AllocationEvent], cfg: Config
) -> list[OutputRow]:
    """构建"调试日志表"行数组（VBA BuildDebugLogRows，19 列，见《调试日志19列规格说明.md》）。

    关闭=不写数据行；简版=仅 is_final_result=True 的最终结果行；详细=全部事件
    （含回溯尝试、撤销等过程事件）。
    """
    if cfg.debug_log_level == DEBUG_LEVEL_OFF:
        return []
    return [
        _map_debug_event_to_row(evt) for evt in events if _should_include_debug_event(evt, cfg)
    ]


def build_run_history_row(
    stats: RunStats,
    cfg: Config,
    dry_run_mode: bool,
    run_time_text: str = "",
    validate_secs: float = 0,
    alloc_secs: float = 0,
    total_secs: float = 0,
    error_code_dist: str = "",
) -> OutputRow:
    """构建"运行历史记录表"单行（VBA BuildRunHistoryRow，20 列）。

    需求 §5.6 的 17 字段 + 末尾 3 个配置快照字段（最大回溯次数/批号比较模式/
    无保质期哨兵值）。第 1 列"运行编号"由 excel_output.append_run_history
    按表内行数自增生成，此处占位空串；第 17 列"备注"留空供文员手动填写。
    """
    return [
        "",
        run_time_text,
        "Dry Run" if dry_run_mode else "Full Run",
        stats.input_return_rows,
        stats.input_inventory_rows,
        stats.input_shipment_count,
        validate_secs,
        alloc_secs,
        total_secs,
        stats.validation_fail_count,
        stats.alloc_success_count,
        stats.alloc_fail_count,
        error_code_dist,
        stats.total_backtrack_count,
        stats.max_group_backtrack,
        cfg.debug_log_level,
        "",
        cfg.max_backtrack_count,
        LOT_MODE_SENSITIVE if cfg.lot_case_sensitive else LOT_MODE_INSENSITIVE,
        cfg.no_expiry_sentinel,
    ]


# -----------------------------------------------------------------------------
# 私有工具函数
# -----------------------------------------------------------------------------


def _normalize_summary_wms_no(wms_order_no: str) -> str:
    """E07 等无 WMS 退单号的场景统一回填 [N/A] 占位符（VBA OB_NormalizeSummaryWmsNo）。"""
    if len(wms_order_no) > 0:
        return wms_order_no
    return NA_PLACEHOLDER


def _is_summary_only_error(error_code: str) -> bool:
    """E08/E11 只进汇总表，不进异常明细（VBA OB_IsSummaryOnlyError）。"""
    return error_code in (ERR_E08, ERR_E11)


def _should_include_debug_event(evt: AllocationEvent, cfg: Config) -> bool:
    """按日志级别决定是否输出该事件（VBA OB_ShouldIncludeDebugEvent）。"""
    if cfg.debug_log_level == DEBUG_LEVEL_SIMPLE:
        return evt.is_final_result
    if cfg.debug_log_level == DEBUG_LEVEL_DETAIL:
        return True
    return False


def _map_debug_event_to_row(evt: AllocationEvent) -> OutputRow:
    """将 AllocationEvent 映射为 19 列 OutputRow（VBA OB_MapDebugEventToRow）。

    is_final_result / is_revoked 为内部过滤标记，不占输出列。
    """
    return [
        evt.shipment_no,
        evt.sku,
        evt.wms_order_no,
        evt.line_no,
        evt.demand_d,
        evt.process_order,
        evt.dynamic_next_min_qty,
        evt.candidate_qc_count,
        evt.excluded_qc_list,
        evt.strategy_used,
        evt.used_qc,
        evt.qc_before,
        evt.qc_after,
        evt.lot_expiry_combo_count,
        evt.is_backtrack_retry,
        evt.backtrack_no,
        evt.line_status,
        evt.error_code,
        evt.fail_sub_type,
    ]
