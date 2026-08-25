"""M05 分配前校验（对应 VBA modValidate.bas）。

职责：
1. 接收 normalize 标准化结果，执行需求 §4.1 四层校验（E01~E08、E11）。
2. 收集所有命中错误码，不因前层错误跳过后层（除 E08/E11 的跳过规则）：
   - 命中 E04 的物流单号跳过 E08；命中 E04 或 E08 的物流单号跳过 E11。
3. 产出 ValidationResult（含 ValidationIssue 列表），并可通过 build_anomaly_rows
   生成异常明细（仅 E01~E07 可定位的错误进入，E08/E11 不进，见需求 §5.4）。

M04 只判断"某个字段本身是否合法"；本模块把字段问题升级为业务错误码，
并额外检查跨行、跨表、跨 SKU 的规则（行号连续性、数量一致性等）。
"""

from __future__ import annotations

from .models import (
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
    NA_PLACEHOLDER,
    SOURCE_INVENTORY_TABLE,
    SOURCE_RETURN_TABLE,
    AnomalyRow,
    Config,
    FieldNormalizeIssue,
    NormalizedInventoryLine,
    NormalizedReturnLine,
    ValidationIssue,
    ValidationResult,
)

# -----------------------------------------------------------------------------
# 公开函数
# -----------------------------------------------------------------------------


def validate_pre(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
    field_issues: list[FieldNormalizeIssue],
    cfg: Config,
) -> ValidationResult:
    """执行 §4.1 四层校验，返回汇总结果；所有命中问题放入 result.issues。

    cfg 为统一校验接口的预留参数；当前 E01~E11 规则不读取配置值。
    """
    del cfg
    issues: list[ValidationIssue] = []
    failed_shipments: set[str] = set()
    shipment_has_e04: set[str] = set()
    shipment_has_e08: set[str] = set()

    # 第1层：E01~E05（字段级）+ E02（行号重复/不连续）
    _apply_layer1_field_issues(orders, inventory, field_issues, issues, failed_shipments, shipment_has_e04)
    _apply_layer1_line_no_continuity(orders, issues, failed_shipments)

    # 第2层：E06、E07（物流单号集合比对）
    _apply_layer2_shipment_consistency(orders, inventory, issues, failed_shipments)

    # 第3层：E08（数量一致性）；已命中 E04 或任一输入表缺少物流单号时跳过
    _apply_layer3_qty_consistency(orders, inventory, issues, failed_shipments, shipment_has_e04, shipment_has_e08)

    # 第4层：E11（碎片库存）；已命中 E04 或 E08 的物流单号跳过
    _apply_layer4_fragment_inventory(orders, inventory, issues, failed_shipments, shipment_has_e04, shipment_has_e08)

    return ValidationResult(
        has_failures=len(failed_shipments) > 0,
        failed_shipment_count=len(failed_shipments),
        issues=issues,
    )


def build_anomaly_rows(validation_issues: list[ValidationIssue]) -> list[AnomalyRow]:
    """需求 §5.4：仅 E01~E07 可定位的错误进入异常明细；E08/E11 为汇总计算结果，不进入。"""
    return [
        AnomalyRow(
            source_table=issue.source_table,
            excel_row_num=issue.excel_row_num,
            shipment_no=issue.shipment_no,
            wms_order_no=issue.wms_order_no,
            sku=issue.sku,
            field_name=issue.field_name,
            raw_value=issue.raw_value,
            error_code=issue.error_code,
            reason=issue.reason,
        )
        for issue in validation_issues
        if _is_anomaly_detail_error(issue.error_code)
    ]


# -----------------------------------------------------------------------------
# 第1层：字段合法性 + E02
# -----------------------------------------------------------------------------


def _apply_layer1_field_issues(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
    field_issues: list[FieldNormalizeIssue],
    issues: list[ValidationIssue],
    failed_shipments: set[str],
    shipment_has_e04: set[str],
) -> None:
    for field_issue in field_issues:
        _append_field_validation_issue(field_issue, orders, inventory, issues, failed_shipments, shipment_has_e04)


def _append_field_validation_issue(
    field_issue: FieldNormalizeIssue,
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
    issues: list[ValidationIssue],
    failed_shipments: set[str],
    shipment_has_e04: set[str],
) -> None:
    error_code = _map_field_issue_to_error_code(field_issue)

    issue = ValidationIssue(
        shipment_no="",
        wms_order_no="",
        sku="",
        error_code=error_code,
        source_table=field_issue.source_table,
        excel_row_num=field_issue.excel_row_num,
        field_name=field_issue.field_name,
        raw_value=field_issue.raw_value,
        reason=_build_field_issue_reason(field_issue, error_code),
    )

    _fill_issue_context(issue, orders, inventory)
    _append_validation_issue(issues, issue, failed_shipments)

    if error_code == ERR_E04 and issue.shipment_no != "":
        shipment_has_e04.add(issue.shipment_no)


def _map_field_issue_to_error_code(field_issue: FieldNormalizeIssue) -> str:
    if field_issue.issue_kind == ISSUE_KIND_EMPTY:
        return ERR_E01
    if field_issue.field_name == "QC情况":
        return ERR_E03
    if field_issue.field_name == "效期":
        return ERR_E05
    if field_issue.field_name == "数量":
        return ERR_E04
    return ERR_E01


def _build_field_issue_reason(field_issue: FieldNormalizeIssue, error_code: str) -> str:
    if error_code == ERR_E01:
        if field_issue.issue_kind == ISSUE_KIND_EMPTY:
            return "字段为空"
        if field_issue.field_name == "行号":
            return "行号格式不符（须为五位前导零文本）"
        return "关键字段为空或格式异常"
    if error_code == ERR_E03:
        return "QC情况非法（仅允许ZP/QC/NG）"
    if error_code == ERR_E04:
        return "数量非法（非正整数）"
    if error_code == ERR_E05:
        return "效期无法解析为合法日期"
    return "关键字段为空或格式异常"


def _fill_issue_context(
    issue: ValidationIssue,
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
) -> None:
    """按来源表+Excel 行号回填物流单号/WMS退单号/SKU；无法定位或为空时统一填 [N/A]。"""
    if issue.source_table == SOURCE_RETURN_TABLE:
        for row in orders:
            if row.excel_row_num == issue.excel_row_num:
                issue.shipment_no = row.shipment_no
                issue.wms_order_no = row.wms_order_no
                issue.sku = row.sku
                break
    elif issue.source_table == SOURCE_INVENTORY_TABLE:
        for row in inventory:
            if row.excel_row_num == issue.excel_row_num:
                issue.shipment_no = row.shipment_no
                issue.wms_order_no = NA_PLACEHOLDER
                issue.sku = row.sku
                break

    if issue.shipment_no == "":
        issue.shipment_no = NA_PLACEHOLDER
    if issue.wms_order_no == "":
        issue.wms_order_no = NA_PLACEHOLDER
    if issue.sku == "":
        issue.sku = NA_PLACEHOLDER


def _apply_layer1_line_no_continuity(
    orders: list[NormalizedReturnLine],
    issues: list[ValidationIssue],
    failed_shipments: set[str],
) -> None:
    """E02：同一 WMS 退单号下，格式合法的行号须构成从 00001 起的完整连续序列。"""
    wms_groups = _group_return_rows_by_wms(orders)

    for wms_key, row_indexes in wms_groups.items():
        # 按行号数值升序排序（保留对应的行下标），等价 VBA SortLongArray
        pairs = sorted((int(orders[idx].line_no), idx) for idx in row_indexes)
        line_nums = [p[0] for p in pairs]
        sorted_indexes = [p[1] for p in pairs]

        if not line_nums:
            continue

        if _detect_line_no_duplicate(line_nums):
            reason = "退单表行号重复"
        else:
            reason = _detect_line_no_discontinuity(line_nums)
            if reason is None:
                continue

        for idx in sorted_indexes:
            row = orders[idx]
            _append_validation_issue(
                issues,
                ValidationIssue(
                    shipment_no=row.shipment_no,
                    wms_order_no=row.wms_order_no,
                    sku=row.sku,
                    error_code=ERR_E02,
                    source_table=SOURCE_RETURN_TABLE,
                    excel_row_num=row.excel_row_num,
                    field_name="行号",
                    # VBA 中 LineNo 为 Long，赋给 String 型 RawValue 时不带前导零（如 1 → "1"），此处保持一致
                    raw_value=str(int(row.line_no)),
                    reason=reason,
                ),
                failed_shipments,
            )


def _group_return_rows_by_wms(orders: list[NormalizedReturnLine]) -> dict[str, list[int]]:
    """按 WMS 退单号分组（仅行号格式合法且 WMS 非空的行参与），值为行下标列表。"""
    groups: dict[str, list[int]] = {}
    for idx, row in enumerate(orders):
        if row.line_no_valid and row.wms_order_no != "":
            groups.setdefault(row.wms_order_no, []).append(idx)
    return groups


def _detect_line_no_duplicate(line_nums: list[int]) -> bool:
    """入参为已排序行号，相邻相等即重复。"""
    return any(line_nums[i] == line_nums[i + 1] for i in range(len(line_nums) - 1))


def _detect_line_no_discontinuity(line_nums: list[int]) -> str | None:
    """入参为已排序且无重复的行号；返回不连续原因文案，连续时返回 None。"""
    min_val = line_nums[0]
    max_val = line_nums[-1]

    if min_val != 1:
        return "行号不从 00001 起：当前序列首行为 " + _format_line_no(min_val)

    if max_val - min_val + 1 != len(line_nums):
        return _build_gap_reason(line_nums)

    for i in range(len(line_nums) - 1):
        if line_nums[i + 1] != line_nums[i] + 1:
            return _build_gap_reason(line_nums)

    return None


def _build_gap_reason(line_nums: list[int]) -> str:
    text_list = "、".join(_format_line_no(n) for n in line_nums)
    return "行号不连续：当前序列为 " + text_list


def _format_line_no(line_no: int) -> str:
    """仅用于 E02 错误原因的序列展示，不回写输入数据，也不代表系统自动补零。"""
    return f"{line_no:05d}"[-5:]


# -----------------------------------------------------------------------------
# 第2层：E06、E07
# -----------------------------------------------------------------------------


def _apply_layer2_shipment_consistency(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
    issues: list[ValidationIssue],
    failed_shipments: set[str],
) -> None:
    order_shipments = _collect_shipment_nos_from_orders(orders)
    inventory_shipments = _collect_shipment_nos_from_inventory(inventory)

    for shipment_no in order_shipments:
        if shipment_no not in inventory_shipments:
            # E06 按退单行逐行生成明细（与 E07 对称），便于在数据异常明细表中定位到具体行
            for row in orders:
                if row.shipment_no == shipment_no:
                    _append_validation_issue(
                        issues,
                        ValidationIssue(
                            shipment_no=row.shipment_no,
                            wms_order_no=row.wms_order_no,
                            sku=row.sku,
                            error_code=ERR_E06,
                            source_table=SOURCE_RETURN_TABLE,
                            excel_row_num=row.excel_row_num,
                            field_name="物流单号",
                            raw_value=row.shipment_no,
                            reason="物流单号仅存在于退单表",
                        ),
                        failed_shipments,
                    )

    for shipment_no in inventory_shipments:
        if shipment_no not in order_shipments:
            for row in inventory:
                if row.shipment_no == shipment_no:
                    _append_validation_issue(
                        issues,
                        ValidationIssue(
                            shipment_no=row.shipment_no,
                            wms_order_no=NA_PLACEHOLDER,
                            sku=row.sku,
                            error_code=ERR_E07,
                            source_table=SOURCE_INVENTORY_TABLE,
                            excel_row_num=row.excel_row_num,
                            field_name="物流单号",
                            raw_value=row.shipment_no,
                            reason="物流单号仅存在于质检库存表",
                        ),
                        failed_shipments,
                    )


def _collect_shipment_nos_from_orders(orders: list[NormalizedReturnLine]) -> set[str]:
    return {row.shipment_no for row in orders if row.shipment_no != ""}


def _collect_shipment_nos_from_inventory(inventory: list[NormalizedInventoryLine]) -> set[str]:
    return {row.shipment_no for row in inventory if row.shipment_no != ""}


# -----------------------------------------------------------------------------
# 第3层：E08
# -----------------------------------------------------------------------------


def _apply_layer3_qty_consistency(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
    issues: list[ValidationIssue],
    failed_shipments: set[str],
    shipment_has_e04: set[str],
    shipment_has_e08: set[str],
) -> None:
    group_keys = _collect_shipment_sku_groups(orders, inventory)

    # E08 只比较"两张表都存在"的物流单号。
    # 仅存在单侧时应由 E06/E07 精确说明，不再叠加没有业务价值的 E08。
    order_shipments = _collect_shipment_nos_from_orders(orders)
    inventory_shipments = _collect_shipment_nos_from_inventory(inventory)

    for shipment_no, sku in group_keys:
        if shipment_no in shipment_has_e04:
            continue
        if shipment_no not in order_shipments:
            continue
        if shipment_no not in inventory_shipments:
            continue

        order_qty = _sum_order_qty(orders, shipment_no, sku)
        inventory_qty = _sum_inventory_qty(inventory, shipment_no, sku)

        if order_qty != inventory_qty:
            _append_validation_issue(
                issues,
                ValidationIssue(
                    shipment_no=shipment_no,
                    wms_order_no=NA_PLACEHOLDER,
                    sku=sku,
                    error_code=ERR_E08,
                    source_table=NA_PLACEHOLDER,
                    excel_row_num=0,
                    field_name="数量",
                    raw_value=f"{order_qty} vs {inventory_qty}",
                    reason="同物流单号+SKU数量不一致",
                ),
                failed_shipments,
            )
            shipment_has_e08.add(shipment_no)


def _collect_shipment_sku_groups(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
) -> set[tuple[str, str]]:
    """两表合并的 (物流单号, SKU) 组集合（仅物流单号与 SKU 均非空的行参与）。"""
    groups: set[tuple[str, str]] = set()
    for row in orders:
        if row.shipment_no != "" and row.sku != "":
            groups.add((row.shipment_no, row.sku))
    for row in inventory:
        if row.shipment_no != "" and row.sku != "":
            groups.add((row.shipment_no, row.sku))
    return groups


def _sum_order_qty(orders: list[NormalizedReturnLine], shipment_no: str, sku: str) -> int:
    return sum(
        row.qty
        for row in orders
        if row.shipment_no == shipment_no and row.sku == sku and row.qty_valid
    )


def _sum_inventory_qty(inventory: list[NormalizedInventoryLine], shipment_no: str, sku: str) -> int:
    return sum(
        row.qty
        for row in inventory
        if row.shipment_no == shipment_no and row.sku == sku and row.qty_valid
    )


# -----------------------------------------------------------------------------
# 第4层：E11
# -----------------------------------------------------------------------------


def _apply_layer4_fragment_inventory(
    orders: list[NormalizedReturnLine],
    inventory: list[NormalizedInventoryLine],
    issues: list[ValidationIssue],
    failed_shipments: set[str],
    shipment_has_e04: set[str],
    shipment_has_e08: set[str],
) -> None:
    group_keys = _collect_order_shipment_sku_groups(orders)

    for shipment_no, sku in group_keys:
        if shipment_no in shipment_has_e04:
            continue
        if shipment_no in shipment_has_e08:
            continue

        group_min_qty = _calc_group_min_qty(orders, shipment_no, sku)
        if group_min_qty <= 0:
            continue

        qc_totals = _sum_inventory_by_qc(inventory, shipment_no, sku)

        for qc, total_qty in qc_totals.items():
            if 0 < total_qty < group_min_qty:
                _append_validation_issue(
                    issues,
                    ValidationIssue(
                        shipment_no=shipment_no,
                        wms_order_no=NA_PLACEHOLDER,
                        sku=sku,
                        error_code=ERR_E11,
                        source_table=NA_PLACEHOLDER,
                        excel_row_num=0,
                        field_name="QC情况",
                        raw_value=f"{qc}:{total_qty}",
                        reason="QC库存碎片无法分配（0 < T < groupMinQty）",
                    ),
                    failed_shipments,
                )


def _collect_order_shipment_sku_groups(orders: list[NormalizedReturnLine]) -> set[tuple[str, str]]:
    """仅退单表的 (物流单号, SKU) 组（物流单号、SKU 非空且数量合法）。"""
    return {
        (row.shipment_no, row.sku)
        for row in orders
        if row.shipment_no != "" and row.sku != "" and row.qty_valid
    }


def _calc_group_min_qty(orders: list[NormalizedReturnLine], shipment_no: str, sku: str) -> int:
    """该 (物流单号+SKU) 组所有数量合法退单行需求量的最小值；无合法行时返回 0。"""
    qtys = [
        row.qty
        for row in orders
        if row.shipment_no == shipment_no and row.sku == sku and row.qty_valid
    ]
    return min(qtys) if qtys else 0


def _sum_inventory_by_qc(
    inventory: list[NormalizedInventoryLine], shipment_no: str, sku: str
) -> dict[str, int]:
    """按 QC 类型汇总库存数量（仅数量与 QC 均合法的行参与）。"""
    totals: dict[str, int] = {}
    for row in inventory:
        if row.shipment_no == shipment_no and row.sku == sku and row.qty_valid and row.qc_valid:
            totals[row.qc] = totals.get(row.qc, 0) + row.qty
    return totals


# -----------------------------------------------------------------------------
# 通用工具
# -----------------------------------------------------------------------------


def _append_validation_issue(
    issues: list[ValidationIssue],
    issue: ValidationIssue,
    failed_shipments: set[str],
) -> None:
    issues.append(issue)
    if issue.shipment_no != "" and issue.shipment_no != NA_PLACEHOLDER:
        failed_shipments.add(issue.shipment_no)


def _is_anomaly_detail_error(error_code: str) -> bool:
    return error_code in (ERR_E01, ERR_E02, ERR_E03, ERR_E04, ERR_E05, ERR_E06, ERR_E07)
