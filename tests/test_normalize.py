"""M04 数据标准化单元测试（对应 UT-Normalize、TC-46/47/48/49）。"""

from datetime import datetime

import pytest

from returned_inventory.config import build_default_config
from returned_inventory.models import (
    CELL_KIND_BLANK,
    CELL_KIND_EXCEL_DATE,
    CELL_KIND_OTHER,
    CELL_KIND_TEXT,
    ISSUE_KIND_EMPTY,
    ISSUE_KIND_FORMAT_ERROR,
    ISSUE_KIND_RANGE_ERROR,
)
from returned_inventory.normalize import (
    normalize_expiry,
    normalize_line_no,
    normalize_lot_no,
    normalize_qc,
    normalize_qty,
    normalize_required_text,
    normalize_text_date,
)


class TestLineNo:
    """需求 §2.1：恰好 5 位且全为数字字符才合法，不自动补零。"""

    def test_valid_five_digit_text(self):
        assert normalize_line_no("00001") == ("00001", True, "")

    def test_numeric_cell_is_invalid(self):
        # Excel 数值单元格 1 → CStr 等价 "1" → 不满足 5 位 → FormatError
        result, valid, kind = normalize_line_no(1)
        assert (valid, kind) == (False, ISSUE_KIND_FORMAT_ERROR)
        assert result == "1"

    def test_float_cell_is_invalid(self):
        # Excel 数值单元格常读为 float；VBA CStr(1.0)="1"，Python 侧需等价
        assert normalize_line_no(1.0) == ("1", False, ISSUE_KIND_FORMAT_ERROR)

    def test_non_digit_text(self):
        assert normalize_line_no("0001A") == ("0001A", False, ISSUE_KIND_FORMAT_ERROR)

    def test_empty(self):
        assert normalize_line_no(None) == ("", False, ISSUE_KIND_EMPTY)

    def test_whitespace_trimmed(self):
        assert normalize_line_no(" 00001 ") == ("00001", True, "")


class TestQC:
    """TC-47：strip → 大写 → 校验。"""

    @pytest.mark.parametrize("raw", ["ZP", " zp ", "qc", "Ng"])
    def test_valid(self, raw):
        result, valid, _ = normalize_qc(raw)
        assert valid and result in ("ZP", "QC", "NG")

    def test_invalid_value(self):
        assert normalize_qc("QM") == ("QM", False, ISSUE_KIND_FORMAT_ERROR)

    def test_empty(self):
        assert normalize_qc("") == ("", False, ISSUE_KIND_EMPTY)


class TestLotNo:
    """TC-46：批号默认 Trim+大写；敏感模式保留原样。"""

    def test_insensitive_default(self):
        cfg = build_default_config()
        assert normalize_lot_no(" a01 ", cfg) == ("A01", True, "")

    def test_leading_zero_kept(self):
        cfg = build_default_config()
        assert normalize_lot_no("00123", cfg) == ("00123", True, "")

    def test_sensitive_mode(self):
        cfg = build_default_config()
        cfg.lot_case_sensitive = True
        assert normalize_lot_no("a01", cfg) == ("a01", True, "")

    def test_empty(self):
        assert normalize_lot_no(None, build_default_config()) == ("", False, ISSUE_KIND_EMPTY)


class TestQty:
    def test_valid(self):
        assert normalize_qty(5) == (5, True, "")
        assert normalize_qty("12") == (12, True, "")

    def test_non_numeric(self):
        assert normalize_qty("abc") == (0, False, ISSUE_KIND_FORMAT_ERROR)

    @pytest.mark.parametrize("raw", ["-3", "0", "12.9"])
    def test_not_positive_int(self, raw):
        assert normalize_qty(raw) == (0, False, ISSUE_KIND_RANGE_ERROR)

    def test_empty(self):
        assert normalize_qty(None) == (0, False, ISSUE_KIND_EMPTY)


class TestTextDate:
    """TC-49：效期字符串级校验（YYYY/MM/DD 或 YYYY-MM-DD，含闰年）。"""

    @pytest.mark.parametrize(
        "text,expected",
        [
            ("2029/01/01", "2029/01/01"),
            ("2029-01-01", "2029/01/01"),
            ("2028/02/29", "2028/02/29"),  # 闰年
            ("2000/02/29", "2000/02/29"),  # 400 年闰
        ],
    )
    def test_valid(self, text, expected):
        assert normalize_text_date(text) == (expected, True)

    @pytest.mark.parametrize(
        "text",
        [
            "2029/02/29",  # 非闰年
            "1900/02/29",  # 100 年不闰
            "2029/13/01",  # 月份 13
            "2029/00/10",
            "2029/04/31",  # 小月 31 日
            "abc",
            "2029/1/1",  # 段长不符
            "1899/01/01",  # 年份越界
            "2029-01/01",  # 混合分隔符段数不符
            "",
        ],
    )
    def test_invalid(self, text):
        assert normalize_text_date(text) == ("", False)


class TestExpiry:
    """效期三分支（对应 VBA VarType 语义）。"""

    def test_excel_date(self):
        assert normalize_expiry(datetime(2029, 1, 1), CELL_KIND_EXCEL_DATE) == (
            "2029/01/01",
            True,
            "",
        )

    def test_text_value(self):
        assert normalize_expiry("2029-12-31", CELL_KIND_TEXT) == ("2029/12/31", True, "")

    def test_text_invalid(self):
        assert normalize_expiry("2029/13/01", CELL_KIND_TEXT) == ("", False, ISSUE_KIND_FORMAT_ERROR)

    def test_blank(self):
        assert normalize_expiry(None, CELL_KIND_BLANK) == ("", False, ISSUE_KIND_EMPTY)

    def test_other(self):
        assert normalize_expiry(123, CELL_KIND_OTHER) == ("", False, ISSUE_KIND_FORMAT_ERROR)


class TestRequiredText:
    def test_normal(self):
        assert normalize_required_text(" SF123 ") == ("SF123", True, "")

    def test_empty(self):
        assert normalize_required_text(None) == ("", False, ISSUE_KIND_EMPTY)
