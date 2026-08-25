"""M02 配置管理（对应 VBA modConfig.bas）。

职责：
1. 从「输入_配置」工作表读取配置（生产键值结构：参数名 | 值 | 说明）。
2. 对缺失的可选配置使用默认值。
3. 对非法配置尽早报错（ConfigError），避免分配算法在错误参数下运行。
"""

from __future__ import annotations

from typing import Any

from .models import (
    DEBUG_LEVEL_DETAIL,
    DEBUG_LEVEL_OFF,
    DEBUG_LEVEL_SIMPLE,
    DEFAULT_DEBUG_LOG_LEVEL,
    DEFAULT_DETAILED_LOG_LIMIT,
    DEFAULT_LOT_MODE,
    DEFAULT_MAX_BACKTRACK_COUNT,
    DEFAULT_NO_EXPIRY_SENTINEL,
    LOT_MODE_INSENSITIVE,
    LOT_MODE_SENSITIVE,
    Config,
)

HEADER_PARAM_NAME = "参数名"
HEADER_PARAM_VALUE = "值"
HEADER_MAX_BACKTRACK = "最大回溯次数"
HEADER_DEBUG_LEVEL = "调试日志级别"
HEADER_LOT_MODE = "批号比较模式"
HEADER_NO_EXPIRY_SENTINEL = "无保质期哨兵值"
HEADER_DETAILED_LOG_LIMIT = "详细日志单表上限"


class ConfigError(Exception):
    """配置非法或配置表结构异常时抛出。"""


def build_default_config() -> Config:
    """默认配置集中在这里，测试也可以直接调用。"""
    return Config()


def load_config(ws: Any) -> Config:
    """从 openpyxl 工作表读取生产键值结构配置（参数名 | 值 | 说明）。

    工作表为空或只有表头时返回默认配置；非法值抛 ConfigError。
    """
    if ws is None:
        raise ConfigError("读取配置失败：工作表对象为空。")

    name_col = _find_header_column(ws, HEADER_PARAM_NAME)
    value_col = _find_header_column(ws, HEADER_PARAM_VALUE)
    if name_col is None or value_col is None:
        # 无键值表头（例如空配置表）→ 全部默认值
        return build_default_config()

    values: dict[str, Any] = {}
    for row in ws.iter_rows(min_row=2):
        param_name = _trim_text(row[name_col - 1].value)
        if param_name:
            values[param_name] = row[value_col - 1].value

    cfg = build_default_config()
    cfg.max_backtrack_count = _parse_positive_int(
        values.get(HEADER_MAX_BACKTRACK), DEFAULT_MAX_BACKTRACK_COUNT, HEADER_MAX_BACKTRACK
    )
    cfg.debug_log_level = _parse_debug_log_level(
        values.get(HEADER_DEBUG_LEVEL), DEFAULT_DEBUG_LOG_LEVEL
    )
    cfg.lot_case_sensitive = _parse_lot_case_sensitive(
        values.get(HEADER_LOT_MODE), DEFAULT_LOT_MODE
    )
    cfg.no_expiry_sentinel = _parse_required_text(
        values.get(HEADER_NO_EXPIRY_SENTINEL), DEFAULT_NO_EXPIRY_SENTINEL, HEADER_NO_EXPIRY_SENTINEL
    )
    cfg.detailed_log_limit = _parse_positive_int(
        values.get(HEADER_DETAILED_LOG_LIMIT), DEFAULT_DETAILED_LOG_LIMIT, HEADER_DETAILED_LOG_LIMIT
    )
    return cfg


# -----------------------------------------------------------------------------
# 配置值解析
# -----------------------------------------------------------------------------


def _trim_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _parse_positive_int(raw_value: Any, default_value: int, field_name: str) -> int:
    text = _trim_text(raw_value)
    if not text:
        return default_value
    try:
        number = float(text)
    except ValueError:
        raise ConfigError(f"配置项 [{field_name}] 必须是正整数，当前值=[{text}]。") from None
    if not number.is_integer() or number <= 0:
        raise ConfigError(f"配置项 [{field_name}] 必须是正整数，当前值=[{text}]。")
    return int(number)


def _parse_debug_log_level(raw_value: Any, default_value: str) -> str:
    text = _trim_text(raw_value)
    if not text:
        return default_value
    if text in (DEBUG_LEVEL_OFF, DEBUG_LEVEL_SIMPLE, DEBUG_LEVEL_DETAIL):
        return text
    raise ConfigError(f"配置项 [{HEADER_DEBUG_LEVEL}] 只能是 关闭/简版/详细，当前值=[{text}]。")


def _parse_lot_case_sensitive(raw_value: Any, default_value: str) -> bool:
    text = _trim_text(raw_value)
    if not text:
        text = default_value
    if text == LOT_MODE_INSENSITIVE:
        return False
    if text == LOT_MODE_SENSITIVE:
        return True
    raise ConfigError(f"配置项 [{HEADER_LOT_MODE}] 只能是 不敏感/敏感，当前值=[{text}]。")


def _parse_required_text(raw_value: Any, default_value: str, field_name: str) -> str:
    del field_name  # 与 VBA 签名保持一致；当前实现空值回退默认值，不报错
    text = _trim_text(raw_value)
    return text if text else default_value


# -----------------------------------------------------------------------------
# 表头工具
# -----------------------------------------------------------------------------


def _find_header_column(ws: Any, header_name: str) -> int | None:
    """返回 1-based 列号；找不到返回 None。"""
    for cell in ws[1]:
        if _trim_text(cell.value) == header_name:
            return cell.column
    return None
