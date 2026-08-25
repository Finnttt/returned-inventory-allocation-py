"""M12 分配后校验（对应 VBA modPostValidate.bas，需求 §4.3）。

职责：对已经成功分配的物流单号做最后一道一致性检查：
  1. 每个退单行的分配数量合计必须等于退单数量（POST_QTY_MISMATCH）。
  2. 同一个退单行不能同时使用两种 QC（POST_QC_MISMATCH）。
  3. 成功明细中的关键字段必须能回到原始退单行（POST_DATA_MISMATCH）。
  4. 退单号状态必须等于其下所有行状态的聚合结果（POST_STATUS_MISMATCH）。
  5. 已经整单回滚的物流单号不参与成功后校验。

与 M11 的依赖关系：VBA 中本模块读取 M11 产出的 FinalResult(Dictionary)，
键为 DetailCount / Detail_i_*（AllocationDetail 字段 + WMSOrderStatus）与
SummaryCount / Summary_i_*。Python 移植为最小依赖设计，不 import status：
- details：具备 AllocationDetail 全部字段、外加 wms_order_status 字段的
  对象序列（对应 FinalResult 的 Detail_i_*）；
- summaries：WMSStatusEntry 序列（对应 FinalResult 的 Summary_i_*），
  其中 status == STATUS_UNALLOCATED 的条目即整单回滚的物流单号。

与 VBA 的有意偏差：
1. VBA 用 Scripting.Dictionary 承载 PostValidationResult（HasFailures /
   IssueCount / Issue_i_* 平铺键）；Python 用 PostValidationResult dataclass
   持有 PostValidationIssue 列表，issue_count 以只读属性对齐 VBA 的 IssueCount。
2. VBA 本模块只返回结果字典，由编排层决定是否终止；需求 §4.3 要求失败即抛
   E99。Python 提供 assert_post_valid 供 runner 调用：失败时抛 E99Error，
   检查项编码与完整描述（含期望值/实际值文本）放入 context，E99Error 的
   expected/actual 数值字段无对应语义，传 0 占位。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Sequence

from .guards import E99Error
from .models import (
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STATUS_UNALLOCATED,
    NormalizedReturnLine,
    WMSStatusEntry,
)

# 分配后校验问题编码（与 VBA modPostValidate 的私有常量一致）
POST_ERR_QTY_MISMATCH = "POST_QTY_MISMATCH"
POST_ERR_QC_MISMATCH = "POST_QC_MISMATCH"
POST_ERR_DATA_MISMATCH = "POST_DATA_MISMATCH"
POST_ERR_STATUS_MISMATCH = "POST_STATUS_MISMATCH"


@dataclass
class PostValidationIssue:
    """单条分配后校验问题（对应 VBA 的 Issue_i_Code/ShipmentNo/... 一组键）。"""

    code: str  # POST_ERR_*
    shipment_no: str
    wms_order_no: str
    sku: str
    line_no: str
    message: str


@dataclass
class PostValidationResult:
    """分配后校验汇总结果（对应 VBA PostValidationResult 字典）。"""

    has_failures: bool = False
    issues: list[PostValidationIssue] = field(default_factory=list)

    @property
    def issue_count(self) -> int:
        """问题总数（对应 VBA 的 IssueCount 键）。"""
        return len(self.issues)


# -----------------------------------------------------------------------------
# 公开函数
# -----------------------------------------------------------------------------


def validate_post(
    orders: Sequence[NormalizedReturnLine],
    details: Sequence | None,
    summaries: Sequence[WMSStatusEntry] | None,
) -> PostValidationResult:
    """对已分配的物流单号做分配后总账复核（对应 VBA ValidatePost）。

    - orders：标准化退单行（M02 输出）；
    - details：成功分配明细（M11 输出），需具备 AllocationDetail 全部字段
      外加 wms_order_status；无明细时传 None 或空序列；
    - summaries：退单号汇总状态（M11 输出的 WMSStatusEntry 序列），
      其中 status == STATUS_UNALLOCATED 的物流单号视为整单回滚。

    发现的所有问题都记录进返回值（VBA 语义：本模块不改分配结果，只读标签）；
    需要按需求 §4.3 中止运行时，由调用方将返回值交给 assert_post_valid。
    """
    result = PostValidationResult()

    rollback_shipments = _collect_rollback_shipments(summaries)
    summary_status_by_wms = _collect_summary_status(summaries)

    qty_by_line: dict[str, int] = {}
    qc_by_line: dict[str, str] = {}
    qc_conflict_by_line: dict[str, str] = {}
    order_qty_by_line: dict[str, int] = {}

    _collect_order_facts(orders, rollback_shipments, order_qty_by_line)
    _collect_detail_facts(details, qty_by_line, qc_by_line, qc_conflict_by_line)
    _check_orders(orders, rollback_shipments, qty_by_line, qc_conflict_by_line, result)
    _check_detail_integrity(
        details, rollback_shipments, order_qty_by_line, summary_status_by_wms, result
    )
    _check_wms_status(details, rollback_shipments, summary_status_by_wms, result)

    return result


def assert_post_valid(result: PostValidationResult) -> None:
    """分配后校验失败时抛 E99Error（需求 §4.3「立即抛出 E99，停止本次运行」）。

    VBA 中本模块只返回 PostValidationResult，由编排层决定是否 RaiseE99；
    Python 由 runner 在校验后调用本函数。E99Error 的 expected/actual 为
    数值语义，后校验失败原因多为文本，故将检查项编码与完整描述放入
    context，expected/actual 传 0 占位（真实期望值/实际值见 message）。
    """
    if not result.has_failures:
        return
    first = result.issues[0]
    raise E99Error(
        first.shipment_no,
        first.sku,
        0,
        0,
        f"{first.code}：{first.message}（共 {result.issue_count} 项失败）",
    )


# -----------------------------------------------------------------------------
# 核心校验流程
# -----------------------------------------------------------------------------


def _check_orders(
    orders: Sequence[NormalizedReturnLine],
    rollback_shipments: set[str],
    qty_by_line: dict[str, int],
    qc_conflict_by_line: dict[str, str],
    result: PostValidationResult,
) -> None:
    """逐退单行检查数量守恒与 QC 一致性（对应 VBA PV_CheckOrders）。"""
    for order in orders:
        if order.shipment_no in rollback_shipments:
            continue

        line_key = _build_line_key(
            order.shipment_no, order.wms_order_no, order.sku, order.line_no
        )
        actual_qty = qty_by_line.get(line_key, 0)

        # 数量守恒是最基础的后校验：输入要几件，最终成功明细就必须合计几件。
        if actual_qty != order.qty:
            _append_issue(
                result,
                POST_ERR_QTY_MISMATCH,
                order.shipment_no,
                order.wms_order_no,
                order.sku,
                order.line_no,
                f"分配量合计 {actual_qty}，退单量 {order.qty}",
            )

        # 同一退单行允许拆批号/效期，但不能跨 QC，否则后续人工处理含义会变得不一致。
        if line_key in qc_conflict_by_line:
            _append_issue(
                result,
                POST_ERR_QC_MISMATCH,
                order.shipment_no,
                order.wms_order_no,
                order.sku,
                order.line_no,
                f"同一退单行使用了多种 QC：{qc_conflict_by_line[line_key]}",
            )


def _collect_order_facts(
    orders: Sequence[NormalizedReturnLine],
    rollback_shipments: set[str],
    order_qty_by_line: dict[str, int],
) -> None:
    """登记非回滚退单行的退单数量（对应 VBA PV_CollectOrderFacts）。"""
    for order in orders:
        if order.shipment_no not in rollback_shipments:
            order_qty_by_line[
                _build_line_key(
                    order.shipment_no, order.wms_order_no, order.sku, order.line_no
                )
            ] = order.qty


def _collect_detail_facts(
    details: Sequence | None,
    qty_by_line: dict[str, int],
    qc_by_line: dict[str, str],
    qc_conflict_by_line: dict[str, str],
) -> None:
    """汇总成功明细的分配合计与 QC 使用情况（对应 VBA PV_CollectDetailFacts）。"""
    for detail in details or ():
        line_key = _build_line_key(
            detail.shipment_no, detail.wms_order_no, detail.sku, detail.line_no
        )
        qty_by_line[line_key] = qty_by_line.get(line_key, 0) + detail.alloc_qty
        if detail.alloc_qty > 0:
            _record_line_qc(line_key, detail.qc, qc_by_line, qc_conflict_by_line)


def _check_detail_integrity(
    details: Sequence | None,
    rollback_shipments: set[str],
    order_qty_by_line: dict[str, int],
    summary_status_by_wms: dict[str, str],
    result: PostValidationResult,
) -> None:
    """逐条成功明细检查数据完整性（对应 VBA PV_CheckDetailIntegrity）。"""
    for detail in details or ():
        ship_no = detail.shipment_no
        wms_order_no = detail.wms_order_no
        sku = detail.sku
        line_no = detail.line_no

        if ship_no in rollback_shipments:
            _append_issue(
                result,
                POST_ERR_DATA_MISMATCH,
                ship_no,
                wms_order_no,
                sku,
                line_no,
                "整单回滚物流单号不应出现在成功分配明细中",
            )
            continue

        if _is_blank(ship_no) or _is_blank(wms_order_no) or _is_blank(sku) or _is_blank(line_no):
            _append_issue(
                result,
                POST_ERR_DATA_MISMATCH,
                ship_no,
                wms_order_no,
                sku,
                line_no,
                "成功分配明细关键字段为空",
            )
            continue

        line_key = _build_line_key(ship_no, wms_order_no, sku, line_no)
        if line_key not in order_qty_by_line:
            _append_issue(
                result,
                POST_ERR_DATA_MISMATCH,
                ship_no,
                wms_order_no,
                sku,
                line_no,
                "成功分配明细找不到对应退单行",
            )
            continue

        expected_order_qty = order_qty_by_line[line_key]
        if detail.order_qty != expected_order_qty:
            _append_issue(
                result,
                POST_ERR_DATA_MISMATCH,
                ship_no,
                wms_order_no,
                sku,
                line_no,
                f"成功明细退单数量 {detail.order_qty}，输入退单数量 {expected_order_qty}",
            )

        if detail.alloc_qty <= 0:
            _append_issue(
                result,
                POST_ERR_DATA_MISMATCH,
                ship_no,
                wms_order_no,
                sku,
                line_no,
                "成功分配明细分配数量必须大于 0",
            )

        if _is_blank(detail.qc) or _is_blank(detail.lot_no) or _is_blank(detail.expiry):
            _append_issue(
                result,
                POST_ERR_DATA_MISMATCH,
                ship_no,
                wms_order_no,
                sku,
                line_no,
                "成功分配明细 QC/批号/效期不能为空",
            )

        detail_wms_status = detail.wms_order_status
        if wms_order_no in summary_status_by_wms:
            if detail_wms_status != summary_status_by_wms[wms_order_no]:
                _append_issue(
                    result,
                    POST_ERR_STATUS_MISMATCH,
                    ship_no,
                    wms_order_no,
                    sku,
                    line_no,
                    f"明细退单号状态 {detail_wms_status}，"
                    f"汇总退单号状态 {summary_status_by_wms[wms_order_no]}",
                )


def _check_wms_status(
    details: Sequence | None,
    rollback_shipments: set[str],
    summary_status_by_wms: dict[str, str],
    result: PostValidationResult,
) -> None:
    """检查退单号状态是否等于行状态聚合结果（对应 VBA PV_CheckWMSStatus）。

    聚合规则（需求 §4.4 第二层）：任一行状态为「手工操作」则退单号状态为
    「手工操作」，否则为「批量导入」。
    """
    expected_by_wms: dict[str, str] = {}
    sample_ship_by_wms: dict[str, str] = {}
    sample_sku_by_wms: dict[str, str] = {}
    sample_line_by_wms: dict[str, str] = {}

    for detail in details or ():
        ship_no = detail.shipment_no
        if ship_no in rollback_shipments:
            continue

        wms_order_no = detail.wms_order_no
        if _is_blank(wms_order_no):
            continue

        if wms_order_no not in sample_ship_by_wms:
            sample_ship_by_wms[wms_order_no] = ship_no
            sample_sku_by_wms[wms_order_no] = detail.sku
            sample_line_by_wms[wms_order_no] = detail.line_no

        line_status = detail.line_status
        if line_status == STATUS_MANUAL:
            expected_by_wms[wms_order_no] = STATUS_MANUAL
        elif line_status == STATUS_BATCH_IMPORT:
            if wms_order_no not in expected_by_wms:
                expected_by_wms[wms_order_no] = STATUS_BATCH_IMPORT
        else:
            _append_issue(
                result,
                POST_ERR_STATUS_MISMATCH,
                ship_no,
                wms_order_no,
                detail.sku,
                detail.line_no,
                f"行状态非法：{line_status}",
            )

    for wms_order_no, expected_status in expected_by_wms.items():
        if wms_order_no not in summary_status_by_wms:
            _append_issue(
                result,
                POST_ERR_STATUS_MISMATCH,
                sample_ship_by_wms[wms_order_no],
                wms_order_no,
                sample_sku_by_wms[wms_order_no],
                sample_line_by_wms[wms_order_no],
                "汇总表缺少该 WMS 退单号状态",
            )
        elif summary_status_by_wms[wms_order_no] != expected_status:
            _append_issue(
                result,
                POST_ERR_STATUS_MISMATCH,
                sample_ship_by_wms[wms_order_no],
                wms_order_no,
                sample_sku_by_wms[wms_order_no],
                sample_line_by_wms[wms_order_no],
                f"行状态聚合应为 {expected_status}，"
                f"汇总表实际为 {summary_status_by_wms[wms_order_no]}",
            )


def _record_line_qc(
    line_key: str,
    qc_value: str,
    qc_by_line: dict[str, str],
    qc_conflict_by_line: dict[str, str],
) -> None:
    """登记退单行使用的 QC，发现第二种 QC 时记录冲突（对应 VBA PV_RecordLineQc）。"""
    if len(qc_value) == 0:
        return
    if line_key not in qc_by_line:
        qc_by_line[line_key] = qc_value
        return
    if qc_by_line[line_key] != qc_value:
        qc_conflict_by_line[line_key] = f"{qc_by_line[line_key]},{qc_value}"


# -----------------------------------------------------------------------------
# 整单回滚跳过规则
# -----------------------------------------------------------------------------


def _collect_rollback_shipments(
    summaries: Sequence[WMSStatusEntry] | None,
) -> set[str]:
    """汇总 status == STATUS_UNALLOCATED 的物流单号（对应 VBA PV_CollectRollbackShipments）。"""
    return {
        entry.shipment_no
        for entry in summaries or ()
        if entry.status == STATUS_UNALLOCATED
    }


def _collect_summary_status(
    summaries: Sequence[WMSStatusEntry] | None,
) -> dict[str, str]:
    """汇总表 WMS 退单号 → 状态映射（对应 VBA PV_CollectSummaryStatus）。"""
    return {
        entry.wms_order_no: entry.status
        for entry in summaries or ()
        if len(entry.wms_order_no) > 0
    }


# -----------------------------------------------------------------------------
# 结果构造与通用取值
# -----------------------------------------------------------------------------


def _append_issue(
    result: PostValidationResult,
    issue_code: str,
    ship_no: str,
    wms_order_no: str,
    sku: str,
    line_no: str,
    message: str,
) -> None:
    """追加一条校验问题（对应 VBA PV_AppendIssue / PV_AppendDetailIssue）。"""
    result.issues.append(
        PostValidationIssue(
            code=issue_code,
            shipment_no=ship_no,
            wms_order_no=wms_order_no,
            sku=sku,
            line_no=line_no,
            message=message,
        )
    )
    result.has_failures = True


def _build_line_key(ship_no: str, wms_order_no: str, sku: str, line_no: str) -> str:
    """退单行唯一键（对应 VBA PV_BuildLineKey，Tab 分隔）。"""
    return f"{ship_no}\t{wms_order_no}\t{sku}\t{line_no}"


def _is_blank(value: str) -> bool:
    """空白判定（对应 VBA PV_IsBlank，含首尾空格视为空）。"""
    return len(value.strip()) == 0
