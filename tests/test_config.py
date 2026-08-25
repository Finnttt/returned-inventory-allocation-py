"""M02 配置管理单元测试（对应 TC-54a/54b 配置非法值场景）。"""

import pytest
from openpyxl import Workbook

from returned_inventory.config import ConfigError, build_default_config, load_config
from returned_inventory.models import (
    DEBUG_LEVEL_DETAIL,
    DEFAULT_DETAILED_LOG_LIMIT,
    DEFAULT_MAX_BACKTRACK_COUNT,
    DEFAULT_NO_EXPIRY_SENTINEL,
)


def make_config_ws(rows):
    wb = Workbook()
    ws = wb.active
    ws.append(["参数名", "值", "说明"])
    for name, value in rows:
        ws.append([name, value, ""])
    return ws


class TestDefaultConfig:
    def test_defaults(self):
        cfg = build_default_config()
        assert cfg.max_backtrack_count == DEFAULT_MAX_BACKTRACK_COUNT == 200
        assert cfg.debug_log_level == "关闭"
        assert cfg.detailed_log_limit == DEFAULT_DETAILED_LOG_LIMIT == 100000
        assert cfg.lot_case_sensitive is False
        assert cfg.no_expiry_sentinel == DEFAULT_NO_EXPIRY_SENTINEL == "2099/01/01"

    def test_empty_sheet_uses_defaults(self):
        wb = Workbook()
        cfg = load_config(wb.active)
        assert cfg == build_default_config()

    def test_header_only_uses_defaults(self):
        cfg = load_config(make_config_ws([]))
        assert cfg == build_default_config()


class TestKeyValueParsing:
    def test_all_values(self):
        ws = make_config_ws(
            [
                ("最大回溯次数", 50),
                ("调试日志级别", "详细"),
                ("批号比较模式", "敏感"),
                ("无保质期哨兵值", "9999/12/31"),
                ("详细日志单表上限", 5000),
            ]
        )
        cfg = load_config(ws)
        assert cfg.max_backtrack_count == 50
        assert cfg.debug_log_level == DEBUG_LEVEL_DETAIL
        assert cfg.lot_case_sensitive is True
        assert cfg.no_expiry_sentinel == "9999/12/31"
        assert cfg.detailed_log_limit == 5000

    def test_partial_values_fall_back_to_defaults(self):
        ws = make_config_ws([("最大回溯次数", 10)])
        cfg = load_config(ws)
        assert cfg.max_backtrack_count == 10
        assert cfg.debug_log_level == "关闭"
        assert cfg.no_expiry_sentinel == "2099/01/01"

    def test_whitespace_tolerated(self):
        ws = make_config_ws([(" 最大回溯次数 ", " 30 ")])
        assert load_config(ws).max_backtrack_count == 30


class TestInvalidValues:
    @pytest.mark.parametrize("bad", ["abc", "-5", "0", "1.5"])
    def test_max_backtrack_must_be_positive_int(self, bad):
        ws = make_config_ws([("最大回溯次数", bad)])
        with pytest.raises(ConfigError, match="最大回溯次数"):
            load_config(ws)

    def test_debug_level_must_be_valid(self):
        ws = make_config_ws([("调试日志级别", "全量")])
        with pytest.raises(ConfigError, match="关闭/简版/详细"):
            load_config(ws)

    def test_lot_mode_must_be_valid(self):
        ws = make_config_ws([("批号比较模式", "大小写")])
        with pytest.raises(ConfigError, match="不敏感/敏感"):
            load_config(ws)

    def test_none_sheet_raises(self):
        with pytest.raises(ConfigError):
            load_config(None)
