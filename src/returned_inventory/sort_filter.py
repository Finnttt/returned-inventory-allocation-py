"""M07 排序·预检测·QC筛选（对应 VBA modSortFilter.bas）。

职责：分配前的全部准备工作。三个接口严格分离，禁止共享隐藏状态：

1. build_static_plan —— 静态排序 + 每行初始可用QC数（需求 §4.2.1，R071）。
   排序在分配开始前一次性完成，分配过程中不重新排序。
2. run_precheck —— 预检测A/B 快速失败检测（需求 §4.2.3，R081/R082），只读不改。
3. filter_candidate_pool —— 每行分配前动态筛选候选QC池（需求 §4.2.2，R072）。

nextMinQty 定义（需求 §4.2.6）：当前行分配后剩余未处理行需求量的最小值（不含当前行）。
静态阶段所有行均未处理，"剩余未处理行"即"同组其他所有行"，两个接口由此共用同一定义。
最后一行 nextMinQty 不存在（is_last_row=True），可用规则收紧为仅 T=D。

可用QC判定（§4.2.2 R072，T=该QC当前库存总量，D=本行需求）：
- 非最后行：T = D（精确匹配无碎片）或 T >= D + nextMinQty（碎片能被后续行消化）
- 最后行：仅 T = D（精确耗尽）
- T <= 0 时直接判为不可用

与 VBA 的有意偏差：
1. VBA 的 StaticPlan 是 Scripting.Dictionary（"Qty_1"/"LineKey_1" 等扁平键）；
   Python 用强类型 dataclass（StaticPlan.lines 元组），承载的信息完全等价。
2. VBA 用稳定冒泡排序；Python 用 sorted（同为稳定排序），四级比较键逐一等价。
   WMS退单号/行号均按字符串字典序比较，不做数值转换（行号为五位前导零文本，
   字符串升序与数值升序结果等价，对应需求 §4.2.1 末排序说明）。
3. 预检测B：VBA 代码实现为 ``supply < forcedDemand``（仅覆盖 T < S）；
   本移植按需求 §4.2.3 与 §6.3.1.2 伪代码实现完整判定
   ``T != S 且 T < S + minQtyOther``（无非强制行时 minQtyOther 视为 +∞，
   即只要 T != S 即命中），额外覆盖 S < T < S + minQtyOther 的碎片场景。
4. VBA 的 filter_candidate_pool 因定长数组需两遍扫描（先统计再填充）；
   Python 用 list 单遍收集，返回结果一致。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from .ledger import InventoryLedger
from .models import (
    QC_PRIORITY,
    CandidateRow,
    NormalizedReturnLine,
    PrecheckResult,
)

# QC 遍历顺序即 models.QC_PRIORITY 的平局兜底顺序（ZP=1 → QC=2 → NG=3）。
# build_static_plan / run_precheck / filter_candidate_pool 三处统一使用，
# 保证"初始可用QC数"与"动态候选池"的 QC 枚举口径一致。
_QC_ORDER: tuple[str, ...] = tuple(sorted(QC_PRIORITY, key=QC_PRIORITY.__getitem__))


# -----------------------------------------------------------------------------
# 公开数据结构
# -----------------------------------------------------------------------------


@dataclass(frozen=True)
class PlanLine:
    """StaticPlan 中的单行（lines 已按四级规则排序完成）。

    init_qc_count：初始可用QC种类数（静态快照，分配开始前一次性计算）。
    """

    wms_order_no: str
    line_no: str
    qty: int
    excel_row_num: int  # Excel 源行号，供 M09 错误定位使用
    init_qc_count: int

    @property
    def line_key(self) -> str:
        """复合键 "WMSOrderNo:LineNo"，供 filter_candidate_pool 唯一定位。"""
        return make_plan_line_key(self.wms_order_no, self.line_no)


@dataclass(frozen=True)
class StaticPlan:
    """静态分配计划（对应 VBA BuildStaticPlan 返回的 Dictionary）。

    空输入时返回各字段为默认值的空计划（row_count=0），不抛错，
    对应 VBA 返回不含 "RowCount" 键的空 Dictionary。
    """

    shipment_no: str = ""
    sku: str = ""
    group_min_qty: int = 0  # 组内最小需求量（E11 校验中已算出，此处冗余保存）
    lines: tuple[PlanLine, ...] = ()

    @property
    def row_count(self) -> int:
        return len(self.lines)


# -----------------------------------------------------------------------------
# 公开函数
# -----------------------------------------------------------------------------


def build_static_plan(
    rows: list[NormalizedReturnLine], ledger: InventoryLedger
) -> StaticPlan:
    """构建静态分配计划（一次性调用，分配开始前执行；调用时账本尚未做任何扣减）。

    步骤：计算 groupMinQty → 每行初始 nextMinQty（=同组其他行需求最小值）与
    初始可用QC数 → 按四级规则稳定排序（可用QC数升序 → Qty降序 →
    WMS退单号升序 → 行号升序）→ 写入 StaticPlan。
    整组只有 1 行时该行等同"最后行"，可用条件收紧为仅 T=D。
    """
    if not rows:
        return StaticPlan()

    ship_no = rows[0].shipment_no
    sku = rows[0].sku
    qtys = [r.qty for r in rows]
    n = len(rows)
    group_min_qty = min(qtys)
    is_only_row = n == 1

    init_qc_counts = [
        _calc_available_qc_count(
            ship_no, sku, qtys[i], _calc_next_min_qty_all(i, qtys), is_only_row, ledger
        )
        for i in range(n)
    ]

    # 四级排序键：(init_qc_count 升序, qty 降序, wms_order_no 升序, line_no 升序)。
    # sorted 为稳定排序，与 VBA 稳定冒泡排序结果一致。
    order = sorted(
        range(n),
        key=lambda i: (
            init_qc_counts[i],
            -qtys[i],
            rows[i].wms_order_no,
            rows[i].line_no,
        ),
    )

    lines = tuple(
        PlanLine(
            wms_order_no=rows[i].wms_order_no,
            line_no=rows[i].line_no,
            qty=rows[i].qty,
            excel_row_num=rows[i].excel_row_num,
            init_qc_count=init_qc_counts[i],
        )
        for i in order
    )
    return StaticPlan(
        shipment_no=ship_no, sku=sku, group_min_qty=group_min_qty, lines=lines
    )


def run_precheck(plan: StaticPlan | None, ledger: InventoryLedger) -> PrecheckResult:
    """执行预检测A / B（只读，不修改 plan 或 ledger）。任一命中立即返回。

    预检测A（§4.2.3 R081）：排序后某行 init_qc_count = 0 → 该行必然无法分配，E09。
    预检测B（§4.2.3 R082）：初始可用QC数=1 且锁定同一QC 的行 >= 2（强制竞争），
    设其合计需求 S、该QC总库存 T、非强制行（可用该QC但可用QC数>1）的最小需求
    minQtyOther（不存在时视为 +∞）；若 T != S 且 T < S + minQtyOther 则命中，E09。
    """
    result = PrecheckResult()
    if plan is None or plan.row_count == 0:
        return result

    n = plan.row_count
    is_only_row = n == 1  # 与 build_static_plan 保持一致：仅 n=1 时是"最后行模式"
    qtys = [line.qty for line in plan.lines]

    # === 预检测A：任意行 init_qc_count = 0 ===
    for line in plan.lines:
        if line.init_qc_count == 0:
            result.precheck_a_hit = True
            return result

    # === 预检测B：多行强制竞争同一QC，合计需求超出该QC可供应量 ===
    for qc in _QC_ORDER:
        t = ledger.query_qc_total(plan.shipment_no, plan.sku, qc)
        forced_demand = 0
        forced_count = 0
        min_qty_other: int | None = None

        for pos in range(n):
            line = plan.lines[pos]
            # 重算该行在静态计划中的 nextMinQty（与 build_static_plan 相同公式：
            # 排序只改变行的次序，不改变"其他所有行"这一集合）
            nmq = _calc_next_min_qty_all(pos, qtys)
            if not _is_qc_available_by_rule(t, line.qty, nmq, is_only_row):
                continue
            if line.init_qc_count == 1:
                # 锁定行：该QC是其唯一可用QC，需求计入强制竞争总量
                forced_demand += line.qty
                forced_count += 1
            elif min_qty_other is None or line.qty < min_qty_other:
                # 非强制行：也可用该QC，其最小需求决定碎片能否被消化
                min_qty_other = line.qty

        if forced_count >= 2 and t != forced_demand:
            # min_qty_other 为 None（无非强制行）时视为 +∞：T != S 即命中
            if min_qty_other is None or t < forced_demand + min_qty_other:
                result.precheck_b_hit = True
                return result

    return result


def filter_candidate_pool(
    current_line_key: str,
    plan: StaticPlan | None,
    ledger: InventoryLedger,
    tried_qcs: Iterable[str] = (),
) -> list[CandidateRow]:
    """动态筛选当前行的候选QC池（每行分配前调用一次）。

    按 current_line_key 在 plan 中定位当前行位置 p，动态 nextMinQty =
    排序位置在 p 之后（尚未分配）的行需求最小值；p 之后无行则为最后行，
    可用规则收紧为仅 T=D。ledger 为实时状态（已反映本 SKU 组前几行的扣减）。

    返回统一候选池（不按 QC 分组，策略一/二/三均在其中操作）：
    按 _QC_ORDER 顺序拼接各可用QC的五元组行，过滤 current_qty=0 的行，
    并排除 tried_qcs 中已尝试失败的QC。无候选、行键不存在或 plan 为空时
    返回空列表（对应 VBA 返回未初始化空数组，调用方用长度判断）。

    注意：本函数不依赖 build_static_plan 的任何缓存状态，只使用 plan 中的
    排序信息；nextMinQty 定义与 build_static_plan 完全相同（参数来源不同）。
    """
    if plan is None or plan.row_count == 0:
        return []

    # --- 定位当前行在排序计划中的位置 ---
    current_pos: int | None = None
    for pos, line in enumerate(plan.lines):
        if line.line_key == current_line_key:
            current_pos = pos
            break
    if current_pos is None:
        # 找不到当前行（调用参数有误），安全返回空
        return []

    d = plan.lines[current_pos].qty

    # --- 动态计算 nextMinQty：当前行之后所有未处理行的最小需求 ---
    qtys_after = [line.qty for line in plan.lines[current_pos + 1 :]]
    is_last_row = not qtys_after
    next_min_qty = min(qtys_after) if qtys_after else 0

    tried = set(tried_qcs)  # 大小写敏感，业务数据已标准化为大写

    pool: list[CandidateRow] = []
    for qc in _QC_ORDER:
        if qc in tried:
            continue
        t = ledger.query_qc_total(plan.shipment_no, plan.sku, qc)
        if not _is_qc_available_by_rule(t, d, next_min_qty, is_last_row):
            continue
        for row in ledger.get_five_tuple_rows(plan.shipment_no, plan.sku, qc):
            if row.current_qty > 0:
                pool.append(row)
    return pool


def make_plan_line_key(wms_order_no: str, line_no: str) -> str:
    """构建 filter_candidate_pool 所需的行唯一标识符（复合键）。

    同一物流单号+SKU下，不同WMS退单号可能有相同行号（如00001），
    单独用行号无法唯一标识一行。分隔符用冒号：业务字段中均不含冒号。
    """
    return f"{wms_order_no}:{line_no}"


# -----------------------------------------------------------------------------
# 私有辅助函数
# -----------------------------------------------------------------------------


def _is_qc_available_by_rule(
    t: int, d: int, next_min_qty: int, is_last_row: bool
) -> bool:
    """QC可用判定规则（§4.2.2 R072）。

    build_static_plan 与 filter_candidate_pool 共用本函数，确保两者使用
    完全相同的 nextMinQty 定义（相同判定逻辑，参数来源不同）。
    """
    if t <= 0:
        return False
    if is_last_row:
        # 最后一行：库存必须恰好等于需求，不能多也不能少
        return t == d
    # 非最后行：精确匹配，或扣除需求后仍能覆盖后续行的最小需求
    return t == d or t >= d + next_min_qty


def _calc_available_qc_count(
    ship_no: str,
    sku: str,
    d: int,
    next_min_qty: int,
    is_last_row: bool,
    ledger: InventoryLedger,
) -> int:
    """计算某行的可用QC种类数：遍历 {ZP, QC, NG} 统计满足可用规则的种类数。"""
    return sum(
        1
        for qc in _QC_ORDER
        if _is_qc_available_by_rule(
            ledger.query_qc_total(ship_no, sku, qc), d, next_min_qty, is_last_row
        )
    )


def _calc_next_min_qty_all(pos: int, qtys: list[int]) -> int:
    """nextMinQty 静态快照：除 pos 行外其他所有行需求量的最小值。

    静态阶段无行已分配，"当前行分配后剩余未处理行" = "全组其余行"。
    只有 1 行时返回 0（外层以 is_only_row/is_last_row 标记处理，不读取该值）。
    """
    others = [qty for j, qty in enumerate(qtys) if j != pos]
    return min(others) if others else 0
