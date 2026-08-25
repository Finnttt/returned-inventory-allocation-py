"""M08 分配策略（对应 VBA modStrategies.bas，需求 §4.2.4 / §4.2.5 / §6.3.2）。

职责：在候选池内依次尝试三级策略（精确匹配 → 最接近匹配 → 多行拼凑），
返回分配尝试结果（含 undo_log）。只解决"当前行如何从候选池拿到 D 件"
——不做回溯，不记调试日志；行状态不在本模块判定，由 M11 按实际
批号/效期组合数判定，本模块只负责把策略名记入明细。

与 VBA 的有意偏差：
- VBA 的 AllocationAttempt 是 Scripting.Dictionary（Success/StrategyUsed/
  QC_1/LotNo_1/... 平铺键），Python 用 AllocationAttempt dataclass +
  AllocationAttemptDetail 列表承载，字段与 VBA 键一一对应。
- 效期比较：VBA 侧已注明"YYYY/MM/DD 字典序与日期序一致"（modStrategies.bas
  CompareByPriority 注释）；Python 直接做字符串比较，标准化层（M04）已保证
  效期统一为 YYYY/MM/DD 格式，二者等价（需求 §4.2.5）。
- used_qc 是 Python 侧新增的派生便利字段（VBA 由 M09 接收结果后自行从明细
  提取）：成功时为实际使用的 QC（策略三即锁定 QC），失败时为空串。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from functools import cmp_to_key

from .ledger import InventoryLedger, UndoLog, new_undo_log
from .models import (
    QC_PRIORITY,
    STRATEGY_FAILED,
    STRATEGY_ONE,
    STRATEGY_THREE,
    STRATEGY_TWO,
    CandidateRow,
)

# 未知 QC 的优先级兜底值（对应 VBA M08_QCPriorityRank 的 Case Else = 99）。
_UNKNOWN_QC_RANK = 99


# -----------------------------------------------------------------------------
# 分配尝试结果结构（对应 VBA AllocationAttempt 字典）
# -----------------------------------------------------------------------------


@dataclass
class AllocationAttemptDetail:
    """单条分配明细（对应 VBA 的 QC_i / LotNo_i / Expiry_i / AllocQty_i 四个键）。

    ShipmentNo/WMSOrderNo/SKU/LineNo/OrderQty/LineStatus 由 M09 在接收结果后
    补填，M08 在分配时还不知道这些上层业务字段。
    """

    qc: str
    lot_no: str
    expiry: str
    alloc_qty: int


@dataclass
class AllocationAttempt:
    """一次分配尝试的结果（对应 VBA TryAllocate 返回的 Dictionary）。

    success：是否分配成功。
    strategy_used：策略一 / 策略二 / 策略三 / 失败（对应 VBA StrategyUsed）。
    used_qc：实际使用的 QC（策略三即第一步锁定的 QC）；失败时为空串。
    undo_log：本次尝试的全部扣减记录，M09 回溯时交给 ledger.undo；
              失败时为空日志，保证 M09 可无条件调用 undo 而不崩溃。
    details：分配明细列表，成功时 len >= 1，失败时为空（对应 VBA DetailCount）。
    """

    success: bool
    strategy_used: str
    used_qc: str
    undo_log: UndoLog
    details: list[AllocationAttemptDetail] = field(default_factory=list)

    @property
    def detail_count(self) -> int:
        """明细条数（对应 VBA DetailCount 键）。"""
        return len(self.details)


# -----------------------------------------------------------------------------
# 公开函数
# -----------------------------------------------------------------------------


def try_allocate(
    pool: list[CandidateRow], demand: int, ledger: InventoryLedger
) -> AllocationAttempt:
    """从候选池为当前退单行分配 demand 件库存（对应 VBA TryAllocate）。

    执行顺序：策略一 → 策略二 → 策略三，首个成功立即返回。
    全部失败时返回 success=False 且账本不变（各策略内部已保证回滚完整）。
    """
    # 边界检查：需求量无效或候选池为空时无需尝试
    if demand <= 0 or not pool:
        return AllocationAttempt(
            success=False,
            strategy_used=STRATEGY_FAILED,
            used_qc="",
            undo_log=new_undo_log(),
        )

    # 按平局优先级对候选池排序（ZP > QC > NG；相同 QC 则效期降序；再相同则批号升序）
    # 排序目的：策略一/二可直接取排序后的第一个命中；策略三平局比较依赖同一规则
    sorted_pool = _sort_pool_by_priority(pool)

    # 依次尝试三个策略，任一成功即结束
    for strategy in (_strategy_one, _strategy_two, _strategy_three):
        attempt = strategy(sorted_pool, demand, ledger)
        if attempt is not None:
            return attempt

    # 全部失败：返回空 undo_log，确保 M09 可无条件调用 undo 而不崩溃
    return AllocationAttempt(
        success=False,
        strategy_used=STRATEGY_FAILED,
        used_qc="",
        undo_log=new_undo_log(),
    )


def compare_by_priority(a: CandidateRow, b: CandidateRow) -> int:
    """比较两个候选行的排序优先级（对应 VBA CompareByPriority，需求 §6.3.2）。

    返回值约定（与排序语义一致）：
      < 0 → a 排在 b 前（a 优先）
      = 0 → 优先级相同
      > 0 → b 排在 a 前（b 优先）

    三级比较规则：
      1. QC 优先级：ZP > QC > NG（数值小者优先，纯确定性兜底，无业务含义）
      2. 效期降序：同 QC 时效期更晚者优先
      3. 批号字母升序：同 QC 同效期时的平局兜底
    """
    # 规则1：QC 优先级
    ra = _qc_priority_rank(a.qc)
    rb = _qc_priority_rank(b.qc)
    if ra != rb:
        return -1 if ra < rb else 1

    # 规则2：效期降序（YYYY/MM/DD 标准化字符串的字典序与日期顺序一致，见模块头注释）
    if a.expiry != b.expiry:
        return -1 if a.expiry > b.expiry else 1

    # 规则3：批号字母升序（平局兜底，确保排序结果可重复）
    if a.lot_no != b.lot_no:
        return -1 if a.lot_no < b.lot_no else 1

    return 0


# -----------------------------------------------------------------------------
# 私有策略函数（成功返回 AllocationAttempt，失败返回 None 且账本不变）
# -----------------------------------------------------------------------------


def _strategy_one(
    sorted_pool: list[CandidateRow], demand: int, ledger: InventoryLedger
) -> AllocationAttempt | None:
    """策略一：精确匹配（对应 VBA StrategyOne）。

    在已排序的候选池中找 current_qty == demand 的行。
    由于池已按 compare_by_priority 排序，第一个命中的即为最高优先级的精确匹配。
    失败：账本不变（无任何扣减操作）。
    """
    for row in sorted_pool:
        if row.current_qty == demand:
            undo_log = new_undo_log()
            if ledger.deduct(row.key, demand, undo_log):
                return AllocationAttempt(
                    success=True,
                    strategy_used=STRATEGY_ONE,
                    used_qc=row.qc,
                    undo_log=undo_log,
                    details=[
                        AllocationAttemptDetail(row.qc, row.lot_no, row.expiry, demand)
                    ],
                )
    return None


def _strategy_two(
    sorted_pool: list[CandidateRow], demand: int, ledger: InventoryLedger
) -> AllocationAttempt | None:
    """策略二：最小 sufficient 匹配（对应 VBA StrategyTwo）。

    找 current_qty > demand 中数量最小的行——选"最小能覆盖"而非最大的，
    保留较大库存供后续行使用（§剩余库存保留原则）。
    相同 current_qty 时：池已排序（高优先级在前），第一个遇到的即为胜者。
    成功：扣减量 = demand（非行全量），该行剩余库存保留在账本/候选池中。
    失败：无 qty > demand 的行，账本不变。
    """
    best_idx = -1
    best_qty = 0
    for i, row in enumerate(sorted_pool):
        if row.current_qty > demand:
            if best_idx == -1 or row.current_qty < best_qty:
                # 发现更小的 sufficient 行（数量更接近 demand，浪费更少）
                best_idx = i
                best_qty = row.current_qty
            # qty 相同时不更新：已排序的池保证先遇到高优先级，不覆盖

    if best_idx == -1:
        return None

    row = sorted_pool[best_idx]
    undo_log = new_undo_log()
    if ledger.deduct(row.key, demand, undo_log):
        return AllocationAttempt(
            success=True,
            strategy_used=STRATEGY_TWO,
            used_qc=row.qc,
            undo_log=undo_log,
            details=[AllocationAttemptDetail(row.qc, row.lot_no, row.expiry, demand)],
        )
    return None


def _strategy_three(
    sorted_pool: list[CandidateRow], demand: int, ledger: InventoryLedger
) -> AllocationAttempt | None:
    """策略三：同 QC 内按"最接近剩余需求量"逐步拼凑（对应 VBA StrategyThree，严禁跨 QC 合并）。

    第一步：从全部候选行选 |qty-demand| 最小者；平局按 QC→晚效期→小批号。
            选中后锁定该行 QC。
    后续：只在锁定 QC 内选 |qty-remaining| 最小者；距离相同时优先
          qty>=remaining（一步完成），再按晚效期→小批号。
    允许最后一步部分扣减；任一步无法继续时撤销本次全部扣减并返回失败。
    """
    first_idx = _find_strategy_three_first(sorted_pool, demand)
    if first_idx == -1:
        return None

    target_qc = sorted_pool[first_idx].qc
    used = [False] * len(sorted_pool)
    temp_log = new_undo_log()
    remaining = demand
    details: list[AllocationAttemptDetail] = []

    while remaining > 0:
        if not details:
            selected_idx = first_idx
        else:
            selected_idx = _find_strategy_three_next(
                sorted_pool, used, target_qc, remaining
            )

        if selected_idx == -1:
            ledger.undo(temp_log)
            return None

        row = sorted_pool[selected_idx]
        to_deduct = min(row.current_qty, remaining)

        if not ledger.deduct(row.key, to_deduct, temp_log):
            ledger.undo(temp_log)
            return None

        used[selected_idx] = True
        details.append(
            AllocationAttemptDetail(target_qc, row.lot_no, row.expiry, to_deduct)
        )
        remaining -= to_deduct

    return AllocationAttempt(
        success=True,
        strategy_used=STRATEGY_THREE,
        used_qc=target_qc,
        undo_log=temp_log,
        details=details,
    )


# -----------------------------------------------------------------------------
# 私有辅助函数
# -----------------------------------------------------------------------------


def _sort_pool_by_priority(pool: list[CandidateRow]) -> list[CandidateRow]:
    """对候选池按 compare_by_priority 排序，返回排好序的副本（原列表不变）。

    Python sorted 为稳定排序，与 VBA 冒泡排序的稳定性语义一致；
    同一 SKU 组的候选行数量通常极小（< 20），性能无忧。
    """
    return sorted(pool, key=cmp_to_key(compare_by_priority))


def _find_strategy_three_first(pool: list[CandidateRow], demand: int) -> int:
    """策略三第一步：数量距离优先；距离相同才使用统一平局规则。"""
    best_idx = -1
    best_diff = 0
    for i, row in enumerate(pool):
        if row.current_qty > 0:
            this_diff = _quantity_distance(row.current_qty, demand)
            if (
                best_idx == -1
                or this_diff < best_diff
                or (this_diff == best_diff and compare_by_priority(row, pool[best_idx]) < 0)
            ):
                best_idx = i
                best_diff = this_diff
    return best_idx


def _find_strategy_three_next(
    pool: list[CandidateRow],
    used: list[bool],
    target_qc: str,
    remaining: int,
) -> int:
    """策略三后续步骤：锁定 QC 后按距离选行；等距时先选可一次完成者，
    若覆盖能力也相同，再按晚效期、批号小的顺序确定结果。"""
    best_idx = -1
    best_diff = 0
    for i, row in enumerate(pool):
        if not used[i] and row.qc == target_qc and row.current_qty > 0:
            this_diff = _quantity_distance(row.current_qty, remaining)

            choose_this = False
            if best_idx == -1:
                choose_this = True
            elif this_diff < best_diff:
                choose_this = True
            elif this_diff == best_diff:
                this_covers = row.current_qty >= remaining
                best_covers = pool[best_idx].current_qty >= remaining
                if this_covers and not best_covers:
                    choose_this = True
                elif this_covers == best_covers:
                    choose_this = compare_by_priority(row, pool[best_idx]) < 0

            if choose_this:
                best_idx = i
                best_diff = this_diff
    return best_idx


def _quantity_distance(qty: int, target: int) -> int:
    """数量距离（对应 VBA M08_QuantityDistance，等价于 abs(qty - target)）。"""
    return abs(qty - target)


def _qc_priority_rank(qc: str) -> int:
    """QC 优先级数值（数值越小优先级越高，对应 VBA M08_QCPriorityRank）。

    使用 models.QC_PRIORITY（ZP=1, QC=2, NG=3）；VBA 用 0/1/2，
    相对顺序一致，比较结果等价。未知 QC 排到最后，不中断运行。
    """
    return QC_PRIORITY.get(qc, _UNKNOWN_QC_RANK)
