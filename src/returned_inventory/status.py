"""M11 状态判定（对应 VBA modStatus.bas，需求 §4.4 / §5.1）。

职责：接收 M09 分配结果与 M05 校验结果，执行：
1. 行级状态判定（批号+效期组合数 → 批量导入 / 手工操作，§4.4.1）
2. 物流单号级整单回滚（任一 SKU 组失败则丢弃该单全部成功明细，§4.4.2）
3. 退单号状态聚合（手工操作优先；失败单 → 无法分配 + 原因字段，§5.1）

原因字段两种格式（与 VBA 逐字一致，供对拍）：
- 直接原因：`{错误码} - {原因说明}`，多码按 E01→E99 升序、同码去重、`; ` 合并
- 连带回滚：`整单回滚（触发原因：{错误码}）`

与 VBA 的结构差异：VBA 的 FinalResult 为 Scripting.Dictionary 平铺键
（SummaryCount / Summary_i_* / DetailCount / Detail_i_*），Python 用
FinalResult dataclass 承载，summary_entries 与 VBA Summary_i_* 一一对应，
details 与 Detail_i_* 一一对应（Detail_i_WMSOrderStatus 对应
FinalDetailRow.wms_order_status）。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .models import (
    ERR_E01,
    ERR_E02,
    ERR_E03,
    ERR_E04,
    ERR_E05,
    ERR_E06,
    ERR_E07,
    ERR_E08,
    ERR_E09,
    ERR_E10,
    ERR_E11,
    ERR_E99,
    LINE_STATUS_FAILED,
    NA_PLACEHOLDER,
    SOURCE_RETURN_TABLE,
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STATUS_UNALLOCATED,
    AllocationDetail,
    NormalizedReturnLine,
    ShipmentAllocResult,
    ValidationIssue,
    ValidationResult,
    WMSStatusEntry,
)

# 错误码 → 标准原因说明（对应 VBA ST_GetStandardReasonText，文案逐字一致）
_STANDARD_REASON_TEXT = {
    ERR_E01: "关键字段为空或格式异常",
    ERR_E02: "退单表行号重复或不连续",
    ERR_E03: "QC情况非法",
    ERR_E04: "数量非法",
    ERR_E05: "效期格式非法",
    ERR_E06: "物流单号仅存在于退单表",
    ERR_E07: "物流单号仅存在于质检库存表",
    ERR_E08: "同物流单号+SKU数量不一致",
    ERR_E09: "分配路径穷尽",
    ERR_E10: "回溯超限",
    ERR_E11: "QC库存碎片无法分配",
    ERR_E99: "未知异常",
}

# 以 (物流单号+SKU) 为粒度触发、需区分直接原因/连带回滚的错误码
# （对应 VBA ST_IsSplitReasonErrorCode）
_SPLIT_REASON_ERROR_CODES = (ERR_E08, ERR_E09, ERR_E10, ERR_E11, ERR_E99)

# 分配阶段直接错误码（对应 VBA ST_IsDirectAllocErrorCode）；
# 连带回滚组（error_code=连带回滚）不属于直接原因
_DIRECT_ALLOC_ERROR_CODES = (ERR_E09, ERR_E10, ERR_E99)


# -----------------------------------------------------------------------------
# 1. 输出结构（对应 VBA FinalResult Dictionary）
# -----------------------------------------------------------------------------


@dataclass
class FinalDetailRow:
    """整单成功物流单号的成功分配明细行（AllocationDetail 字段 + 退单号状态）。

    对应 VBA FinalResult 的 Detail_i_* 键组；wms_order_status 对应
    Detail_i_WMSOrderStatus（同一退单号的所有明细行该值相同，需求 §5.2）。
    """

    shipment_no: str
    wms_order_no: str
    sku: str
    line_no: str
    order_qty: int
    qc: str
    lot_no: str
    expiry: str
    alloc_qty: int
    line_status: str  # 批量导入 / 手工操作 / 分配失败
    wms_order_status: str  # 该行所属 WMS 退单号的聚合状态
    strategy_used: str = ""


@dataclass
class FinalResult:
    """状态判定最终结果（对应 VBA ApplyRollback 返回的 Dictionary）。

    summary_entries：汇总表行（对应 SummaryCount / Summary_i_*），
    含成功（批量导入/手工操作）与失败（无法分配+原因）两类。
    details：成功分配明细（对应 DetailCount / Detail_i_*），
    仅整单成功的物流单号有明细；被整单回滚的物流单号不写明细。
    """

    summary_entries: list[WMSStatusEntry] = field(default_factory=list)
    details: list[FinalDetailRow] = field(default_factory=list)


# -----------------------------------------------------------------------------
# 2. 公开函数
# -----------------------------------------------------------------------------


def determine_line_status(details: list[AllocationDetail], line_no: str) -> str:
    """按某退货行实际使用的"批号+效期"组合种类数判定行状态（VBA DetermineLineStatus）。

    组合数 = 1 → 批量导入；≥ 2 → 手工操作；无有效明细 → 分配失败。
    注意：此公开函数只按 line_no 统计，适合单 WMS/单 SKU 的简单测试。
    生产聚合路径（apply_rollback）使用完整退货行身份
    （物流单号+WMS退单号+SKU+行号），避免不同退单号的 00001 互相污染。
    """
    combo_count = _count_lot_expiry_combos(details, line_no)
    if combo_count == 0:
        return LINE_STATUS_FAILED
    if combo_count == 1:
        return STATUS_BATCH_IMPORT
    return STATUS_MANUAL


def build_rollback_reason(direct_codes: list[str], trigger_code: str) -> str:
    """构造"无法分配"原因字符串（VBA BuildRollbackReason）。

    direct_codes 为空 → 连带回滚格式；非空 → 按 E01→E99 升序去重后拼接直接原因。
    """
    if len(direct_codes) == 0:
        return f"整单回滚（触发原因：{trigger_code}）"
    return _format_direct_codes(direct_codes)


def apply_rollback(
    shipment_results: list[ShipmentAllocResult],
    validation_result: ValidationResult,
    validation_issues: list[ValidationIssue],
    orders: list[NormalizedReturnLine],
) -> FinalResult:
    """汇总校验失败与分配失败，执行整单回滚，产出 FinalResult（VBA ApplyRollback）。

    某物流单号下存在任何失败行（校验阶段 E01~E11 或分配阶段 E09/E10/E99，
    含连带回滚组）时，该单全部分配结果撤回、所有 WMS 退单号标记"无法分配"。
    orders 可传空列表；用于 E08 等错误把 [N/A] 展开到具体 WMS 退单号。

    validation_result 仅为对齐 VBA 签名保留（VBA ApplyRollback 同样只接收
    不在函数体内使用；M05 的问题明细全部由 validation_issues 承载）。
    """
    del validation_result
    result = FinalResult()

    shipment_nos = _collect_all_shipment_nos(shipment_results, validation_issues)

    for ship_no in shipment_nos:
        alloc = _find_shipment_result(shipment_results, ship_no)
        validation_failed = _shipment_has_validation_issue(validation_issues, ship_no)
        allocation_failed = _shipment_allocation_failed(alloc)

        if validation_failed or allocation_failed:
            # 整单回滚：不写成功明细，只写汇总表"无法分配"行
            _append_failure_summary(result, ship_no, validation_issues, alloc, orders)
        else:
            # 全部 SKU 组成功：写入明细并聚合退单号状态
            _append_success_output(result, ship_no, alloc)

    return result


def aggregate_wms_status(final_result: FinalResult | None) -> list[WMSStatusEntry]:
    """从 FinalResult 中提取汇总表记录（VBA AggregateWMSStatus / WMSStatusMap）。"""
    if final_result is None:
        return []
    return [
        WMSStatusEntry(
            shipment_no=entry.shipment_no,
            wms_order_no=entry.wms_order_no,
            status=entry.status,
            reason=entry.reason,
        )
        for entry in final_result.summary_entries
    ]


# -----------------------------------------------------------------------------
# 3. 整单回滚 / 成功输出
# -----------------------------------------------------------------------------


def _append_failure_summary(
    result: FinalResult,
    ship_no: str,
    validation_issues: list[ValidationIssue],
    alloc: ShipmentAllocResult | None,
    orders: list[NormalizedReturnLine],
) -> None:
    """失败物流单号：每个 WMS 退单号写一行"无法分配"汇总（VBA ST_AppendFailureSummary）。"""
    wms_orders = _collect_wms_orders(ship_no, validation_issues, alloc, orders)
    trigger_code = _find_trigger_code(validation_issues, alloc, ship_no)

    for wms_order_no in wms_orders:
        direct_codes = _collect_direct_codes_for_wms(
            ship_no, wms_order_no, validation_issues, alloc, orders
        )
        reason = build_rollback_reason(direct_codes, trigger_code)
        result.summary_entries.append(
            WMSStatusEntry(
                shipment_no=ship_no,
                wms_order_no=wms_order_no,
                status=STATUS_UNALLOCATED,
                reason=reason,
            )
        )


def _append_success_output(
    result: FinalResult,
    ship_no: str,
    alloc: ShipmentAllocResult | None,
) -> None:
    """成功物流单号：重算行状态、聚合退单号状态，写明细与汇总（VBA ST_AppendSuccessOutput）。"""
    if alloc is None:
        return

    all_details = _extract_details(alloc)
    if len(all_details) == 0:
        return

    # 行状态以"同一退货行实际使用的批号+效期组合数"为准。
    # 注意：不同 WMS 退单号都会有 00001，不能只按行号统计，否则会把不同退货行误合并。
    for detail in all_details:
        detail.line_status = _determine_detail_line_status(all_details, detail)

    # 先聚合每个 WMS 退单号的退单号状态
    wms_status = _build_wms_status_map(all_details)

    # 写成功明细
    for detail in all_details:
        result.details.append(
            FinalDetailRow(
                shipment_no=detail.shipment_no,
                wms_order_no=detail.wms_order_no,
                sku=detail.sku,
                line_no=detail.line_no,
                order_qty=detail.order_qty,
                qc=detail.qc,
                lot_no=detail.lot_no,
                expiry=detail.expiry,
                alloc_qty=detail.alloc_qty,
                line_status=detail.line_status,
                wms_order_status=wms_status[detail.wms_order_no],
                strategy_used=detail.strategy_used,
            )
        )

    # 写成功汇总（每个 WMS 退单号一行，原因为空）
    for wms_order_no, status in wms_status.items():
        result.summary_entries.append(
            WMSStatusEntry(
                shipment_no=ship_no,
                wms_order_no=wms_order_no,
                status=status,
                reason="",
            )
        )


# -----------------------------------------------------------------------------
# 4. 行状态与原因格式
# -----------------------------------------------------------------------------


def _count_lot_expiry_combos(details: list[AllocationDetail], line_no: str) -> int:
    """按行号统计 (批号+效期) 组合数（VBA ST_CountLotExpiryCombos），alloc_qty=0 不计。"""
    combos = {
        (d.lot_no, d.expiry)
        for d in details
        if d.line_no == line_no and d.alloc_qty > 0
    }
    return len(combos)


def _determine_detail_line_status(
    details: list[AllocationDetail], target: AllocationDetail
) -> str:
    """按完整退货行身份判定行状态（VBA ST_DetermineDetailLineStatus）。"""
    combo_count = _count_lot_expiry_combos_for_detail(details, target)
    if combo_count == 0:
        return LINE_STATUS_FAILED
    if combo_count == 1:
        return STATUS_BATCH_IMPORT
    return STATUS_MANUAL


def _count_lot_expiry_combos_for_detail(
    details: list[AllocationDetail], target: AllocationDetail
) -> int:
    """按完整退货行身份（物流单号+WMS退单号+SKU+行号）统计组合数
    （VBA ST_CountLotExpiryCombosForDetail）。"""
    combos = {
        (d.lot_no, d.expiry)
        for d in details
        if _is_same_return_line(d, target) and d.alloc_qty > 0
    }
    return len(combos)


def _is_same_return_line(left: AllocationDetail, right: AllocationDetail) -> bool:
    """完整退货行身份判定（VBA ST_IsSameReturnLine）。"""
    return (
        left.shipment_no == right.shipment_no
        and left.wms_order_no == right.wms_order_no
        and left.sku == right.sku
        and left.line_no == right.line_no
    )


def _format_direct_codes(codes: list[str]) -> str:
    """多错误码升序去重后格式化为 `码 - 说明`，`; ` 合并（VBA ST_FormatDirectCodes）。"""
    sorted_codes = _sort_unique_codes(codes)
    parts = [f"{code} - {_get_standard_reason_text(code)}" for code in sorted_codes]
    return "; ".join(parts)


def _get_standard_reason_text(error_code: str) -> str:
    """错误码标准原因说明（VBA ST_GetStandardReasonText），未知码 → "未知错误"。"""
    return _STANDARD_REASON_TEXT.get(error_code, "未知错误")


def _is_split_reason_error_code(error_code: str) -> bool:
    """是否需区分直接原因/连带回滚的错误码（VBA ST_IsSplitReasonErrorCode）。"""
    return error_code in _SPLIT_REASON_ERROR_CODES


def _is_direct_alloc_error_code(error_code: str) -> bool:
    """是否为分配阶段直接错误码（VBA ST_IsDirectAllocErrorCode）。"""
    return error_code in _DIRECT_ALLOC_ERROR_CODES


# -----------------------------------------------------------------------------
# 5. 物流单号 / WMS 退单号收集
# -----------------------------------------------------------------------------


def _collect_all_shipment_nos(
    shipment_results: list[ShipmentAllocResult],
    validation_issues: list[ValidationIssue],
) -> list[str]:
    """合并分配结果与校验问题中的全部物流单号，保插入序去重（VBA ST_CollectAllShipmentNos）。"""
    seen: dict[str, bool] = {}
    for shipment_result in shipment_results:
        seen[shipment_result.shipment_no] = True
    for issue in validation_issues:
        if issue.shipment_no != "":
            seen[issue.shipment_no] = True
    return list(seen.keys())


def _find_shipment_result(
    shipment_results: list[ShipmentAllocResult], ship_no: str
) -> ShipmentAllocResult | None:
    """按物流单号定位分配结果（VBA ST_FindShipmentResult）。"""
    for shipment_result in shipment_results:
        if shipment_result.shipment_no == ship_no:
            return shipment_result
    return None


def _shipment_has_validation_issue(
    validation_issues: list[ValidationIssue], ship_no: str
) -> bool:
    """该物流单号是否存在任一校验问题（VBA ST_ShipmentHasValidationIssue）。"""
    return any(issue.shipment_no == ship_no for issue in validation_issues)


def _shipment_allocation_failed(alloc: ShipmentAllocResult | None) -> bool:
    """该物流单号是否存在任一失败的 SKU 组（VBA ST_ShipmentAllocationFailed）。

    连带回滚组 success=False，同样视为失败（短路总原则下触发整单回滚）。
    """
    if alloc is None:
        return False
    return any(not group.success for group in alloc.group_results)


def _collect_wms_orders(
    ship_no: str,
    validation_issues: list[ValidationIssue],
    alloc: ShipmentAllocResult | None,
    orders: list[NormalizedReturnLine],
) -> list[str]:
    """收集失败物流单号下需输出汇总行的全部 WMS 退单号，保插入序去重
    （VBA ST_CollectWmsOrders）。"""
    seen: dict[str, bool] = {}

    for issue in validation_issues:
        if issue.shipment_no != ship_no:
            continue
        if issue.wms_order_no != "" and issue.wms_order_no != NA_PLACEHOLDER:
            seen[issue.wms_order_no] = True
        elif (
            issue.wms_order_no == NA_PLACEHOLDER
            and issue.source_table == SOURCE_RETURN_TABLE
            and issue.excel_row_num > 0
        ):
            # 退单表真实数据行的 WMS 为空时，除正常 WMS 汇总外，
            # 还必须保留一条 [N/A]，让文员能看到并定位该异常行。
            # E06 等整单级问题 excel_row_num=0，不应额外生成 [N/A]。
            seen[NA_PLACEHOLDER] = True

    _add_wms_from_alloc(seen, alloc)
    _add_wms_from_orders(seen, ship_no, orders)

    # E07 / E06 / E08 等可能只有 [N/A]：至少保留占位行
    if len(seen) == 0:
        seen[NA_PLACEHOLDER] = True

    return list(seen.keys())


def _add_wms_from_alloc(seen: dict[str, bool], alloc: ShipmentAllocResult | None) -> None:
    """从分配结果明细中补充 WMS 退单号（VBA ST_AddWmsFromAllocMap）。"""
    if alloc is None:
        return
    for group in alloc.group_results:
        for detail in group.details:
            seen[detail.wms_order_no] = True


def _add_wms_from_orders(
    seen: dict[str, bool], ship_no: str, orders: list[NormalizedReturnLine]
) -> None:
    """从退单表标准化行中补充 WMS 退单号（VBA ST_AddWmsFromOrders）。"""
    for order in orders:
        if order.shipment_no == ship_no and order.wms_order_no != "":
            seen[order.wms_order_no] = True


def _find_trigger_code(
    validation_issues: list[ValidationIssue],
    alloc: ShipmentAllocResult | None,
    ship_no: str,
) -> str:
    """整单回滚的触发错误码：取该单全部校验问题与分配直接错误码中排序最前者
    （VBA ST_FindTriggerCode）；均无则兜底 E09。"""
    best_code = ""

    for issue in validation_issues:
        if issue.shipment_no == ship_no:
            best_code = _pick_earlier_error_code(best_code, issue.error_code)

    if alloc is not None:
        for group in alloc.group_results:
            if _is_direct_alloc_error_code(group.error_code):
                best_code = _pick_earlier_error_code(best_code, group.error_code)

    if best_code == "":
        best_code = ERR_E09
    return best_code


def _collect_direct_codes_for_wms(
    ship_no: str,
    wms_order_no: str,
    validation_issues: list[ValidationIssue],
    alloc: ShipmentAllocResult | None,
    orders: list[NormalizedReturnLine],
) -> list[str]:
    """收集可直接归因到该 WMS 退单号的错误码（VBA ST_CollectDirectCodesForWms）。

    非"分原因"码（E01~E07）适用即直接归因；E08/E11 等需该退单号下确实
    含有触发 SKU 才算直接原因，否则走连带回滚格式。
    """
    code_set: dict[str, bool] = {}

    for issue in validation_issues:
        if issue.shipment_no != ship_no:
            continue
        if _validation_issue_applies_to_wms(issue, wms_order_no, orders):
            if not _is_split_reason_error_code(
                issue.error_code
            ) or _validation_issue_applies_to_wms_directly(issue, wms_order_no, orders):
                code_set[issue.error_code] = True

    _add_alloc_direct_codes(code_set, ship_no, wms_order_no, alloc, orders)

    return list(code_set.keys())


def _validation_issue_applies_to_wms(
    issue: ValidationIssue,
    wms_order_no: str,
    orders: list[NormalizedReturnLine],
) -> bool:
    """校验问题是否波及该 WMS 退单号（VBA ST_ValidationIssueAppliesToWms）。"""
    # E07 等孤立物流单号：汇总行 WMS=[N/A]
    if wms_order_no == NA_PLACEHOLDER:
        return issue.wms_order_no == NA_PLACEHOLDER

    if issue.wms_order_no == wms_order_no:
        return True

    # 物流单号+SKU 级错误（WMS 为 [N/A]）：扩展到该 SKU 下所有 WMS 退单号
    if issue.wms_order_no == NA_PLACEHOLDER and issue.sku != NA_PLACEHOLDER:
        return _wms_has_sku(wms_order_no, issue.shipment_no, issue.sku, orders, None)

    # E06 等整单级错误：扩展到该物流单号下所有 WMS 退单号
    if issue.wms_order_no == NA_PLACEHOLDER:
        return _wms_belongs_to_shipment(wms_order_no, issue.shipment_no, orders)

    return False


def _validation_issue_applies_to_wms_directly(
    issue: ValidationIssue,
    wms_order_no: str,
    orders: list[NormalizedReturnLine],
) -> bool:
    """校验问题是否可直接归因到该 WMS 退单号（VBA ST_ValidationIssueAppliesToWmsDirectly）。"""
    # E01~E07 行级/退单号级：issue 已带具体 WMS
    if not _is_split_reason_error_code(issue.error_code):
        return _validation_issue_applies_to_wms(issue, wms_order_no, orders)

    # E08/E11：issue 挂在物流单号+SKU，退单号下含该 SKU 即视为直接原因
    if issue.wms_order_no == NA_PLACEHOLDER and issue.sku != NA_PLACEHOLDER:
        return _wms_has_sku(wms_order_no, issue.shipment_no, issue.sku, orders, None)

    return False


def _add_alloc_direct_codes(
    code_set: dict[str, bool],
    ship_no: str,
    wms_order_no: str,
    alloc: ShipmentAllocResult | None,
    orders: list[NormalizedReturnLine],
) -> None:
    """把分配阶段直接错误码归因到含该失败 SKU 的 WMS 退单号（VBA ST_AddAllocDirectCodes）。"""
    if alloc is None:
        return

    for group in alloc.group_results:
        if not _is_direct_alloc_error_code(group.error_code):
            continue
        if _wms_has_sku(wms_order_no, ship_no, group.sku, orders, alloc):
            code_set[group.error_code] = True


def _wms_has_sku(
    wms_order_no: str,
    ship_no: str,
    sku: str,
    orders: list[NormalizedReturnLine],
    alloc: ShipmentAllocResult | None,
) -> bool:
    """该 WMS 退单号下是否含有指定 SKU（VBA ST_WmsHasSku）。

    优先查退单表标准化行；orders 未命中时，从分配结果明细中反查 WMS+SKU 关系。
    """
    for order in orders:
        if (
            order.shipment_no == ship_no
            and order.wms_order_no == wms_order_no
            and order.sku == sku
        ):
            return True

    # 未传 orders 时，从分配结果中反查 WMS+SKU 关系
    if sku != "" and alloc is not None:
        return _alloc_links_wms_sku(alloc, wms_order_no, sku)

    return False


def _alloc_links_wms_sku(
    alloc: ShipmentAllocResult, wms_order_no: str, sku: str
) -> bool:
    """分配结果明细中是否存在该 WMS+SKU 组合（VBA ST_AllocMapLinksWmsSku）。

    失败组常无明细：该 SKU 组存在且 WMS 在 orders 未知时，由调用方用 orders 判定。
    """
    for group in alloc.group_results:
        if group.sku != sku:
            continue
        for detail in group.details:
            if detail.wms_order_no == wms_order_no:
                return True
    return False


def _wms_belongs_to_shipment(
    wms_order_no: str, ship_no: str, orders: list[NormalizedReturnLine]
) -> bool:
    """该 WMS 退单号是否属于指定物流单号（VBA ST_WmsBelongsToShipment）。"""
    return any(
        order.shipment_no == ship_no and order.wms_order_no == wms_order_no
        for order in orders
    )


# -----------------------------------------------------------------------------
# 6. 成功路径：明细提取与退单号聚合
# -----------------------------------------------------------------------------


def _extract_details(alloc: ShipmentAllocResult | None) -> list[AllocationDetail]:
    """提取该物流单号全部 SKU 组的分配明细，按组序+组内序（VBA ST_ExtractDetailsFromAllocMap）。"""
    if alloc is None:
        return []
    details: list[AllocationDetail] = []
    for group in alloc.group_results:
        details.extend(group.details)
    return details


def _build_wms_status_map(details: list[AllocationDetail]) -> dict[str, str]:
    """按退单号聚合状态：任一行为手工操作 → 手工操作，否则批量导入
    （VBA ST_BuildWmsStatusMap + ST_WmsHasManualLine）。保插入序。"""
    wms_status: dict[str, str] = {}
    for detail in details:
        wms_status.setdefault(detail.wms_order_no, STATUS_BATCH_IMPORT)
        if detail.line_status == STATUS_MANUAL:
            wms_status[detail.wms_order_no] = STATUS_MANUAL
    return wms_status


# -----------------------------------------------------------------------------
# 7. 通用小工具
# -----------------------------------------------------------------------------


def _pick_earlier_error_code(current_code: str, new_code: str) -> str:
    """取排序更靠前的错误码（VBA ST_PickEarlierErrorCode）。"""
    if current_code == "":
        return new_code
    if _error_code_sort_key(new_code) < _error_code_sort_key(current_code):
        return new_code
    return current_code


def _error_code_sort_key(error_code: str) -> int:
    """错误码排序键（VBA ST_ErrorCodeSortKey）：E+数字 → 数值；其余非 E 前缀 → 9999。

    VBA 中 E 前缀但数字部分非法时 CLng 失败、函数保持默认值 0，此处保持一致。
    """
    if len(error_code) >= 2 and error_code[0] == "E":
        try:
            return int(error_code[1:])
        except ValueError:
            return 0
    return 9999


def _sort_unique_codes(codes: list[str]) -> list[str]:
    """错误码去重并按排序键升序（VBA ST_SortUniqueCodes），空码忽略。"""
    unique: dict[str, bool] = {}
    for code in codes:
        if len(code) > 0:
            unique[code] = True
    return sorted(unique.keys(), key=_error_code_sort_key)
