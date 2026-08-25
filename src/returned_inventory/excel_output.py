"""M14 Excel 写入（对应 VBA modExcelOutput.bas，需求 §5 / §6.7.4）。

职责：接收 M13 生成的 OutputRow，负责输出工作表清空、写入、调试日志分表、
运行历史追加。本模块只做"写表动作"，不做任何业务计算。

输出表头常量在此集中定义（VBA 中汇总/明细/异常表头定义在 M15 RN_WriteAllOutput，
调试日志表头定义在 M14 WriteDebugLog；Python 统一收拢到写表层，列顺序与 VBA 逐字一致）。

与 VBA 的结构差异：
- VBA 的"受保护工作表"检查对应 openpyxl 的 ws.protection.sheet；
  命中时抛 OutputError（文案与 VBA Err 1400~1410 系列一致）。
- VBA 用 NumberFormat="@" 整列设置文本格式；openpyxl 无可靠的列级格式，
  移植为逐单元格设置 number_format="@"（同样的防前导零/日期转换目的）。
"""

from __future__ import annotations

from typing import Any

from .models import (
    DEBUG_LEVEL_OFF,
    DEFAULT_DETAILED_LOG_LIMIT,
    Config,
    OutputRow,
)

# 输出工作表名称（与 VBA/modExcelOutput 私有常量一致；运行历史表名由 runner 管理）
SHEET_SUMMARY = "分配状态汇总表"
SHEET_DETAIL = "成功分配明细表"
SHEET_ANOMALY = "数据异常明细表"
SHEET_DEBUG = "调试日志"

# 分配状态汇总表表头（需求 §5.1，4 列；VBA 定义于 modRunner RN_WriteAllOutput）
SUMMARY_HEADERS = ["物流单号", "WMS退单号", "退单号状态", "原因"]

# 成功分配明细表表头（需求 §5.2，11 列）
DETAIL_HEADERS = [
    "物流单号", "WMS退单号", "SKU", "行号", "退单数量", "QC情况",
    "批号", "效期", "分配数量", "行状态", "退单号状态",
]

# 数据异常明细表表头（需求 §5.4，9 列）
ANOMALY_HEADERS = [
    "来源表", "Excel行号", "物流单号", "WMS退单号", "SKU",
    "字段名", "原始值", "错误码", "原因说明",
]

# 调试日志表头（19 列，与 M13 build_debug_log_rows / 调试日志19列规格说明.md 一致）
DEBUG_HEADERS = [
    "物流单号", "SKU", "WMS退单号", "行号", "D", "处理序", "动态nextMinQty",
    "候选QC数", "被排除QC列表", "策略", "分配QC", "分配前QC剩余", "分配后QC剩余",
    "批号/效期组合数", "是否回溯重试", "实际回溯次数", "行状态", "错误码",
    "分配失败子类型",
]

# 运行历史记录表表头（20 列：需求 §5.6 的 17 字段 + 3 个配置快照字段；
# 与 生成生产工作簿.ps1 的表头逐字一致）
RUN_HISTORY_HEADERS = [
    "运行编号", "运行时间", "运行类型", "输入：退单表行数", "输入：质检库存表行数",
    "输入：物流单号数", "校验耗时（秒）", "分配耗时（秒）", "总耗时（秒）",
    "校验失败物流单号数", "分配成功物流单号数", "分配失败物流单号数",
    "错误码分布", "总回溯次数", "最大单组回溯次数", "调试日志级别",
    "备注", "最大回溯次数", "批号比较模式", "无保质期哨兵值",
]

# 需要按文本格式写入的列（VBA EO_ApplyTextFormats）：
# 行号/批号可能有前导零；效期/哨兵值须保留 YYYY/MM/DD 文本；原始值列承载各类录入原值。
_TEXT_FORMAT_HEADERS = {"行号", "效期", "无保质期哨兵值", "原始值", "批号"}


class OutputError(Exception):
    """输出写入阶段的结构性错误（对应 VBA modExcelOutput 的 Err.Raise 1400~1410 系列）。"""


# -----------------------------------------------------------------------------
# 公开函数（与 VBA modExcelOutput 公开函数一一对应）
# -----------------------------------------------------------------------------


def clear_output_sheets(wb: Any, cfg: Config) -> None:
    """清空输出工作表数据区（保留表头），并删除调试日志动态分表（VBA ClearOutputSheets）。

    只清空输出相关表；输入表、配置表、运行历史记录表不在清空范围（需求 §6.7.4）。
    cfg 为对齐 VBA 签名的预留参数，当前清空逻辑不读取配置值。
    """
    del cfg
    if wb is None:
        raise OutputError("工作簿对象为空，无法清空输出表。")

    for sheet_name in (SHEET_SUMMARY, SHEET_DETAIL, SHEET_ANOMALY, SHEET_DEBUG):
        _clear_data_keep_header(_get_sheet(wb, sheet_name))
    _delete_debug_split_sheets(wb)


def write_sheet(ws: Any, rows: list[OutputRow], headers: list[str]) -> None:
    """向指定工作表写入表头和数据（覆盖数据区，保留第 1 行作为表头）（VBA WriteSheet）。"""
    if ws is None:
        raise OutputError("目标工作表为空，无法写入。")
    _ensure_sheet_writable(ws)
    if not headers:
        raise OutputError(f"表头为空，无法写入工作表 [{ws.title}]。")

    _clear_data_keep_header(ws)

    text_format_cols = {
        idx for idx, header in enumerate(headers, start=1) if header in _TEXT_FORMAT_HEADERS
    }

    for col_index, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_index, value=header)
        if col_index in text_format_cols:
            cell.number_format = "@"

    for row_offset, row in enumerate(rows, start=2):
        for col_index in range(1, len(headers) + 1):
            # 行列数不足时补空串（对应 VBA EO_OutputRowsToMatrix 的 vbNullString 兜底）
            value = row[col_index - 1] if col_index <= len(row) else ""
            cell = ws.cell(row=row_offset, column=col_index, value=value)
            if col_index in text_format_cols:
                cell.number_format = "@"


def write_debug_log(wb: Any, rows: list[OutputRow], cfg: Config) -> None:
    """写入调试日志，详细模式超过 detailed_log_limit 时自动分表（VBA WriteDebugLog）。

    分表命名：调试日志（主表）、调试日志_2、调试日志_3 ……顺延编号；
    每次写入前删除既有分表，避免残留（需求 §6.7.4 第 4 条）。
    调试日志关闭时主表仅保留表头，不写任何数据行。
    """
    if wb is None:
        raise OutputError("工作簿对象为空，无法写入调试日志。")

    _delete_debug_split_sheets(wb)

    limit_per_sheet = cfg.detailed_log_limit
    if limit_per_sheet <= 0:
        limit_per_sheet = DEFAULT_DETAILED_LOG_LIMIT

    main_ws = _get_sheet(wb, SHEET_DEBUG)

    if cfg.debug_log_level == DEBUG_LEVEL_OFF or len(rows) == 0:
        write_sheet(main_ws, [], DEBUG_HEADERS)
        return

    for sheet_no, chunk_start in enumerate(range(0, len(rows), limit_per_sheet), start=1):
        chunk = rows[chunk_start : chunk_start + limit_per_sheet]
        if sheet_no == 1:
            target_ws = main_ws
        else:
            target_ws = _get_or_create_sheet(wb, f"{SHEET_DEBUG}_{sheet_no}")
        write_sheet(target_ws, chunk, DEBUG_HEADERS)


def append_run_history(ws: Any, row: OutputRow) -> None:
    """向运行历史记录表追加单行，不覆盖已有记录（VBA AppendRunHistory）。

    第 1 列"运行编号"按表内行数自增生成：表头占第 1 行，数据第 N 行编号 = N - 1。
    第 2 列"运行时间"与第 20 列"无保质期哨兵值"按文本写入，防止 Excel
    把 2026/07/19 10:33:05、2099/01/01 转成本机日期格式。
    """
    if ws is None:
        raise OutputError("运行历史工作表为空，无法追加。")
    _ensure_sheet_writable(ws)
    if not row:
        return

    next_row = max(ws.max_row + 1, 2)

    for col_index, value in enumerate(row, start=1):
        cell = ws.cell(row=next_row, column=col_index, value=value)
        if col_index in (2, 20):
            cell.number_format = "@"

    ws.cell(row=next_row, column=1).value = next_row - 1


# -----------------------------------------------------------------------------
# 私有工具函数
# -----------------------------------------------------------------------------


def _get_sheet(wb: Any, sheet_name: str) -> Any:
    """按名取工作表，缺失时抛 OutputError（对应 VBA Worksheets(name) 下标越界）。"""
    if sheet_name not in wb.sheetnames:
        raise OutputError(f"工作簿缺少工作表 [{sheet_name}]。")
    return wb[sheet_name]


def _clear_data_keep_header(ws: Any) -> None:
    """清空数据区、保留第 1 行表头（VBA EO_ClearDataKeepHeader）。"""
    _ensure_sheet_writable(ws)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)


def _ensure_sheet_writable(ws: Any) -> None:
    """受保护工作表禁止写入（VBA EO_EnsureSheetWritable）。"""
    if ws.protection.sheet:
        raise OutputError(f"工作表 [{ws.title}] 受保护，已中止写入。")


def _delete_debug_split_sheets(wb: Any) -> None:
    """删除全部调试日志分表（调试日志_2、调试日志_3 …）（VBA EO_DeleteDebugSplitSheets）。"""
    for sheet_name in list(wb.sheetnames):
        if _is_debug_split_sheet(sheet_name):
            ws = wb[sheet_name]
            _ensure_sheet_writable(ws)
            del wb[sheet_name]


def _is_debug_split_sheet(sheet_name: str) -> bool:
    """是否为调试日志分表名：调试日志_<纯数字>（VBA EO_IsDebugSplitSheet + EO_IsAllDigits）。"""
    prefix = f"{SHEET_DEBUG}_"
    if not sheet_name.startswith(prefix):
        return False
    suffix = sheet_name[len(prefix):]
    return len(suffix) > 0 and suffix.isdigit()


def _get_or_create_sheet(wb: Any, sheet_name: str) -> Any:
    """取已存在的工作表，不存在则在末尾新建（VBA EO_GetOrCreateSheet）。"""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    return wb.create_sheet(sheet_name)
