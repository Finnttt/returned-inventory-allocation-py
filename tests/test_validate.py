"""M05 分配前校验单元测试（对应 VBA modValidate.bas / 需求 §4.1、§5.4）。

测试数据全部内存构造，不接触 Excel。
"""

from returned_inventory.models import (
    ERR_E01,
    ERR_E02,
    ERR_E03,
    ERR_E04,
    ERR_E05,
    ERR_E06,
    ERR_E07,
    ERR_E08,
    ERR_E11,
    ISSUE_KIND_EMPTY,
    ISSUE_KIND_FORMAT_ERROR,
    ISSUE_KIND_RANGE_ERROR,
    NA_PLACEHOLDER,
    SOURCE_INVENTORY_TABLE,
    SOURCE_RETURN_TABLE,
    AnomalyRow,
    Config,
    FieldNormalizeIssue,
    NormalizedInventoryLine,
    NormalizedReturnLine,
    ValidationIssue,
)
from returned_inventory.validate import build_anomaly_rows, validate_pre

CFG = Config()


def make_order(
    excel_row_num=2,
    shipment_no="SH1",
    wms_order_no="W1",
    sku="SKU1",
    line_no="00001",
    qty=10,
    line_no_valid=True,
    qty_valid=True,
    empty_fields="",
):
    return NormalizedReturnLine(
        excel_row_num=excel_row_num,
        shipment_no=shipment_no,
        wms_order_no=wms_order_no,
        sku=sku,
        line_no=line_no,
        qty=qty,
        line_no_valid=line_no_valid,
        qty_valid=qty_valid,
        empty_fields=empty_fields,
    )


def make_inventory(
    excel_row_num=2,
    shipment_no="SH1",
    sku="SKU1",
    qc="ZP",
    lot_no="LOT1",
    expiry="2030/01/01",
    qty=10,
    qc_valid=True,
    expiry_valid=True,
    qty_valid=True,
    empty_fields="",
):
    return NormalizedInventoryLine(
        excel_row_num=excel_row_num,
        shipment_no=shipment_no,
        sku=sku,
        qc=qc,
        lot_no=lot_no,
        expiry=expiry,
        qty=qty,
        qc_valid=qc_valid,
        expiry_valid=expiry_valid,
        qty_valid=qty_valid,
        empty_fields=empty_fields,
    )


def make_field_issue(
    excel_row_num=2,
    source_table=SOURCE_RETURN_TABLE,
    field_name="SKU",
    raw_value="",
    issue_kind=ISSUE_KIND_EMPTY,
):
    return FieldNormalizeIssue(
        excel_row_num=excel_row_num,
        source_table=source_table,
        field_name=field_name,
        raw_value=raw_value,
        issue_kind=issue_kind,
    )


def error_codes(result):
    return [issue.error_code for issue in result.issues]


class TestAllValid:
    def test_no_issues(self):
        result = validate_pre([make_order()], [make_inventory()], [], CFG)
        assert result.issues == []
        assert result.has_failures is False
        assert result.failed_shipment_count == 0

    def test_empty_inputs(self):
        result = validate_pre([], [], [], CFG)
        assert result.issues == []
        assert result.has_failures is False


class TestE01:
    def test_empty_field(self):
        orders = [make_order(sku="", empty_fields="SKU")]
        # 库存行 SKU 同样为空，避免引入与本次断言无关的 E06/E08
        inventory = [make_inventory(sku="", empty_fields="SKU")]
        issues = [make_field_issue(field_name="SKU", issue_kind=ISSUE_KIND_EMPTY)]
        result = validate_pre(orders, inventory, issues, CFG)

        assert error_codes(result) == [ERR_E01]
        issue = result.issues[0]
        assert issue.reason == "字段为空"
        assert issue.field_name == "SKU"
        # FillIssueContext：按行号回填所在行上下文，SKU 本身为空则填 [N/A]
        assert issue.shipment_no == "SH1"
        assert issue.wms_order_no == "W1"
        assert issue.sku == NA_PLACEHOLDER

    def test_line_no_format(self):
        orders = [make_order(line_no="123", line_no_valid=False)]
        issues = [make_field_issue(field_name="行号", raw_value="123", issue_kind=ISSUE_KIND_FORMAT_ERROR)]
        result = validate_pre(orders, [make_inventory()], issues, CFG)

        assert error_codes(result) == [ERR_E01]
        assert result.issues[0].reason == "行号格式不符（须为五位前导零文本）"
        assert result.issues[0].raw_value == "123"

    def test_other_format_field(self):
        issues = [make_field_issue(field_name="物流单号", raw_value="@@", issue_kind=ISSUE_KIND_FORMAT_ERROR)]
        result = validate_pre([make_order()], [make_inventory()], issues, CFG)

        assert ERR_E01 in error_codes(result)
        e01 = next(i for i in result.issues if i.error_code == ERR_E01)
        assert e01.reason == "关键字段为空或格式异常"

    def test_context_not_found_fills_na(self):
        # 字段问题行号在标准化行中找不到（如整行被丢弃），上下文全部填 [N/A]
        issues = [make_field_issue(excel_row_num=99, field_name="SKU", issue_kind=ISSUE_KIND_EMPTY)]
        result = validate_pre([make_order()], [make_inventory()], issues, CFG)

        issue = result.issues[0]
        assert (issue.shipment_no, issue.wms_order_no, issue.sku) == (
            NA_PLACEHOLDER,
            NA_PLACEHOLDER,
            NA_PLACEHOLDER,
        )
        # [N/A] 不计入失败物流单号
        assert result.has_failures is False


class TestE02:
    def test_duplicate_line_no(self):
        orders = [
            make_order(excel_row_num=2, line_no="00001", qty=5),
            make_order(excel_row_num=3, line_no="00001", qty=5),
        ]
        result = validate_pre(orders, [make_inventory()], [], CFG)

        assert error_codes(result) == [ERR_E02, ERR_E02]
        assert all(i.reason == "退单表行号重复" for i in result.issues)
        assert all(i.field_name == "行号" for i in result.issues)
        # VBA 中 RawValue 来自 Long 型 LineNo，无前导零
        assert all(i.raw_value == "1" for i in result.issues)
        assert {i.excel_row_num for i in result.issues} == {2, 3}

    def test_not_start_from_00001(self):
        orders = [make_order(line_no="00002")]
        result = validate_pre(orders, [make_inventory()], [], CFG)

        assert error_codes(result) == [ERR_E02]
        assert result.issues[0].reason == "行号不从 00001 起：当前序列首行为 00002"

    def test_gap(self):
        orders = [
            make_order(excel_row_num=2, line_no="00001", qty=5),
            make_order(excel_row_num=3, line_no="00003", qty=5),
        ]
        result = validate_pre(orders, [make_inventory()], [], CFG)

        assert error_codes(result) == [ERR_E02, ERR_E02]
        assert result.issues[0].reason == "行号不连续：当前序列为 00001、00003"

    def test_invalid_line_no_excluded_from_continuity(self):
        # 行号格式非法的行不参与 E02 连续性判断（§5.4 跳过规则）
        orders = [
            make_order(excel_row_num=2, line_no="00001"),
            make_order(excel_row_num=3, line_no="abc", line_no_valid=False, empty_fields="行号"),
        ]
        result = validate_pre(orders, [make_inventory()], [], CFG)
        assert ERR_E02 not in error_codes(result)

    def test_grouped_by_wms(self):
        # 不同 WMS 退单号各自独立检查，各自从 00001 起即合法
        orders = [
            make_order(excel_row_num=2, wms_order_no="W1", line_no="00001", qty=5),
            make_order(excel_row_num=3, wms_order_no="W2", line_no="00001", qty=5),
        ]
        result = validate_pre(orders, [make_inventory()], [], CFG)
        assert result.issues == []


class TestE03E04E05:
    def test_e03_qc_invalid(self):
        inventory = [make_inventory(qc="QM", qc_valid=False)]
        issues = [
            make_field_issue(
                source_table=SOURCE_INVENTORY_TABLE,
                field_name="QC情况",
                raw_value="QM",
                issue_kind=ISSUE_KIND_FORMAT_ERROR,
            )
        ]
        result = validate_pre([make_order()], inventory, issues, CFG)

        assert error_codes(result) == [ERR_E03]
        issue = result.issues[0]
        assert issue.reason == "QC情况非法（仅允许ZP/QC/NG）"
        assert issue.raw_value == "QM"
        # 库存行无 WMS 字段，填 [N/A]
        assert (issue.shipment_no, issue.wms_order_no, issue.sku) == ("SH1", NA_PLACEHOLDER, "SKU1")

    def test_e04_qty_invalid(self):
        orders = [make_order(qty=0, qty_valid=False)]
        issues = [make_field_issue(field_name="数量", raw_value="12.9", issue_kind=ISSUE_KIND_RANGE_ERROR)]
        result = validate_pre(orders, [make_inventory()], issues, CFG)

        assert error_codes(result) == [ERR_E04]
        assert result.issues[0].reason == "数量非法（非正整数）"
        assert result.issues[0].raw_value == "12.9"

    def test_e04_empty_qty_is_e01(self):
        # 数量为空归 E01，不触发 E04（需求 §5.4）
        issues = [make_field_issue(field_name="数量", issue_kind=ISSUE_KIND_EMPTY)]
        result = validate_pre([make_order()], [make_inventory()], issues, CFG)
        assert ERR_E01 in error_codes(result)
        assert ERR_E04 not in error_codes(result)

    def test_e05_expiry_invalid(self):
        inventory = [make_inventory(expiry="2029/13/01", expiry_valid=False)]
        issues = [
            make_field_issue(
                source_table=SOURCE_INVENTORY_TABLE,
                field_name="效期",
                raw_value="2029/13/01",
                issue_kind=ISSUE_KIND_FORMAT_ERROR,
            )
        ]
        result = validate_pre([make_order()], inventory, issues, CFG)

        assert error_codes(result) == [ERR_E05]
        assert result.issues[0].reason == "效期无法解析为合法日期"


class TestE06E07:
    def test_e06_per_order_row(self):
        orders = [
            make_order(excel_row_num=2, line_no="00001", qty=5),
            make_order(excel_row_num=3, line_no="00002", qty=5),
        ]
        result = validate_pre(orders, [], [], CFG)

        assert error_codes(result) == [ERR_E06, ERR_E06]
        assert all(i.reason == "物流单号仅存在于退单表" for i in result.issues)
        assert all(i.source_table == SOURCE_RETURN_TABLE for i in result.issues)
        assert all(i.field_name == "物流单号" and i.raw_value == "SH1" for i in result.issues)
        assert result.failed_shipment_count == 1

    def test_e07_per_inventory_row(self):
        inventory = [make_inventory(shipment_no="SH9")]
        result = validate_pre([make_order()], inventory, [], CFG)

        e07 = [i for i in result.issues if i.error_code == ERR_E07]
        assert len(e07) == 1
        assert e07[0].reason == "物流单号仅存在于质检库存表"
        assert e07[0].source_table == SOURCE_INVENTORY_TABLE
        assert e07[0].wms_order_no == NA_PLACEHOLDER
        assert e07[0].shipment_no == "SH9"

    def test_e06_e07_symmetric(self):
        # 两侧各有孤立物流单号时 E06、E07 同时产出
        result = validate_pre([make_order()], [make_inventory(shipment_no="SH9")], [], CFG)
        codes = error_codes(result)
        assert ERR_E06 in codes and ERR_E07 in codes
        assert result.failed_shipment_count == 2


class TestE08:
    def test_qty_mismatch(self):
        result = validate_pre([make_order(qty=10)], [make_inventory(qty=7)], [], CFG)

        assert error_codes(result) == [ERR_E08]
        issue = result.issues[0]
        assert issue.reason == "同物流单号+SKU数量不一致"
        assert issue.raw_value == "10 vs 7"
        assert issue.field_name == "数量"
        # 跨行汇总比较，无法定位到具体行/表
        assert issue.source_table == NA_PLACEHOLDER
        assert issue.excel_row_num == 0
        assert issue.wms_order_no == NA_PLACEHOLDER
        assert result.failed_shipment_count == 1

    def test_multi_row_sum_compared(self):
        # 按 物流单号+SKU 汇总后比对：多行合计相等则不触发
        orders = [
            make_order(excel_row_num=2, line_no="00001", qty=6),
            make_order(excel_row_num=3, line_no="00002", qty=4),
        ]
        inventory = [make_inventory(excel_row_num=2, qty=3), make_inventory(excel_row_num=3, qty=7)]
        result = validate_pre(orders, inventory, [], CFG)
        assert result.issues == []

    def test_skipped_when_shipment_only_on_one_side(self):
        # 单侧存在的物流单号由 E06/E07 说明，不叠加 E08
        result = validate_pre([make_order()], [], [], CFG)
        assert ERR_E08 not in error_codes(result)

    def test_e04_skips_e08_for_that_shipment_only(self):
        orders = [
            make_order(excel_row_num=2, shipment_no="SH1", qty=0, qty_valid=False),
            make_order(excel_row_num=3, shipment_no="SH2", sku="SKU2", qty=5),
        ]
        inventory = [
            make_inventory(excel_row_num=2, shipment_no="SH1", qty=10),
            make_inventory(excel_row_num=3, shipment_no="SH2", sku="SKU2", qty=7),
        ]
        issues = [make_field_issue(excel_row_num=2, field_name="数量", raw_value="0", issue_kind=ISSUE_KIND_RANGE_ERROR)]
        result = validate_pre(orders, inventory, issues, CFG)

        codes = error_codes(result)
        assert ERR_E04 in codes
        # SH1 命中 E04 → 跳过 E08；SH2 数量 5 vs 7 仍正常产出 E08
        e08 = [i for i in result.issues if i.error_code == ERR_E08]
        assert len(e08) == 1 and e08[0].shipment_no == "SH2"


class TestE11:
    def test_fragment_triggers_e11(self):
        orders = [
            make_order(excel_row_num=2, line_no="00001", qty=10),
            make_order(excel_row_num=3, line_no="00002", qty=10),
        ]
        inventory = [
            make_inventory(excel_row_num=2, qc="ZP", qty=18),
            make_inventory(excel_row_num=3, qc="QC", qty=2),
        ]
        result = validate_pre(orders, inventory, [], CFG)

        assert error_codes(result) == [ERR_E11]
        issue = result.issues[0]
        assert issue.reason == "QC库存碎片无法分配（0 < T < groupMinQty）"
        assert issue.raw_value == "QC:2"
        assert issue.field_name == "QC情况"
        assert issue.source_table == NA_PLACEHOLDER
        assert issue.excel_row_num == 0

    def test_t_equal_group_min_not_triggered(self):
        # T == groupMinQty 不属于 0 < T < groupMinQty
        result = validate_pre([make_order(qty=10)], [make_inventory(qty=10)], [], CFG)
        assert ERR_E11 not in error_codes(result)

    def test_t_zero_not_triggered(self):
        # QC 类型总库存为 0（无合法数量行）不参与 E11
        inventory = [
            make_inventory(excel_row_num=2, qc="ZP", qty=10),
            make_inventory(excel_row_num=3, qc="NG", qty=0, qty_valid=False),
        ]
        result = validate_pre([make_order(qty=10)], inventory, [], CFG)
        assert result.issues == []

    def test_qc_invalid_row_excluded(self):
        # QC 非法行（qc_valid=False）不进入任何 QC 汇总桶，但仍计入 E08 总量（只看 qty_valid）。
        # 构造：订单 6+6（groupMinQty=6，合计 12），库存 ZP 10 + QC非法行 2（E08 守恒 12=12）。
        # 若非法行被错误计入 QC 桶，"XX":2 满足 0 < 2 < 6 会触发 E11。
        orders = [
            make_order(excel_row_num=2, line_no="00001", qty=6),
            make_order(excel_row_num=3, line_no="00002", qty=6),
        ]
        inventory = [
            make_inventory(excel_row_num=2, qc="ZP", qty=10),
            make_inventory(excel_row_num=3, qc="XX", qc_valid=False, qty=2),
        ]
        result = validate_pre(orders, inventory, [], CFG)
        assert result.issues == []

    def test_e04_skips_e11(self):
        orders = [make_order(qty=0, qty_valid=False)]
        inventory = [make_inventory(qty=2)]
        issues = [make_field_issue(field_name="数量", raw_value="0", issue_kind=ISSUE_KIND_RANGE_ERROR)]
        result = validate_pre(orders, inventory, issues, CFG)

        assert ERR_E04 in error_codes(result)
        assert ERR_E11 not in error_codes(result)

    def test_e08_skips_e11(self):
        # 数量不一致（E08）且存在碎片时，E11 跳过
        orders = [make_order(qty=10)]
        inventory = [make_inventory(qty=7)]
        result = validate_pre(orders, inventory, [], CFG)

        assert error_codes(result) == [ERR_E08]
        assert ERR_E11 not in error_codes(result)


class TestMultipleErrorCodes:
    def test_coexist(self):
        # 同一物流单号命中 E03 + E08，全部记录（不因前层错误跳过后层）
        orders = [make_order(qty=10)]
        inventory = [make_inventory(qc="XX", qc_valid=False, qty=7)]
        issues = [
            make_field_issue(
                source_table=SOURCE_INVENTORY_TABLE,
                field_name="QC情况",
                raw_value="XX",
                issue_kind=ISSUE_KIND_FORMAT_ERROR,
            )
        ]
        result = validate_pre(orders, inventory, issues, CFG)

        assert set(error_codes(result)) == {ERR_E03, ERR_E08}
        assert result.has_failures is True
        assert result.failed_shipment_count == 1


class TestBuildAnomalyRows:
    def _make_issue(self, error_code):
        return ValidationIssue(
            shipment_no="SH1",
            wms_order_no="W1",
            sku="SKU1",
            error_code=error_code,
            source_table=SOURCE_RETURN_TABLE,
            excel_row_num=2,
            field_name="数量",
            raw_value="abc",
            reason="原因",
        )

    def test_routing(self):
        # E01~E07 进异常明细；E08、E11 不进（需求 §5.4）
        codes = [ERR_E01, ERR_E02, ERR_E03, ERR_E04, ERR_E05, ERR_E06, ERR_E07, ERR_E08, ERR_E11]
        rows = build_anomaly_rows([self._make_issue(c) for c in codes])

        assert [r.error_code for r in rows] == codes[:7]
        assert all(isinstance(r, AnomalyRow) for r in rows)
        # 字段完整复制
        row = rows[0]
        assert (row.shipment_no, row.wms_order_no, row.sku, row.field_name, row.raw_value, row.reason) == (
            "SH1",
            "W1",
            "SKU1",
            "数量",
            "abc",
            "原因",
        )

    def test_empty_input(self):
        assert build_anomaly_rows([]) == []

    def test_only_e08_e11_yields_empty(self):
        rows = build_anomaly_rows([self._make_issue(ERR_E08), self._make_issue(ERR_E11)])
        assert rows == []

    def test_end_to_end_with_validate_pre(self):
        # E08/E11 出现在 issues 中但不进异常明细
        orders = [make_order(qty=10)]
        inventory = [make_inventory(qty=7)]
        result = validate_pre(orders, inventory, [], CFG)

        assert error_codes(result) == [ERR_E08]
        assert build_anomaly_rows(result.issues) == []
