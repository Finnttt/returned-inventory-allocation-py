"""M10 工程守卫（对应 VBA modGuards.bas，需求 §6.5）。

职责：提供运行时强制断言，检验库存账目数学等式成立。
任一断言失败立即抛 E99Error（对应 VBA RaiseE99 的 Err.Raise 自定义错误号），
终止当前分配运行，防止错误结果被静默写入输出。

与 M09 的协作关系（守卫挂载点，需求 §6.3.1.3 / §6.5）：
- 每个 SKU 组分配完成后：assert_conservation（§6.5.1 组结束点守恒断言）；
- 每次回溯撤销后：assert_undo_consistency（§6.5.2 撤销点检验）；
- E10 全量回滚后：再次 assert_undo_consistency（对应 VBA E10 分支的断言，
  此时选择栈必为空，退化为"账本与快照精确相等"）。

与 VBA 的有意偏差：
1. VBA 的 AssertConservation / AssertUndoConsistency 返回 Boolean，由调用方
   （M09）再调用 RaiseE99 抛错；Python 将两步融合——断言失败时直接抛 E99Error，
   成功时返回 True。净行为等价（VBA 中断言为 False 后必然紧跟 RaiseE99），
   少一层中转，且错误对象直接携带物流单号/SKU/期望/实际/上下文五个字段。
2. VBA 的 AssertUndoConsistency 只处理"撤销后账本与快照完全相等"
   （choiceStack 为预留参数，源码中传 Nothing 也可工作）；Python 按需求 §6.5.2
   实现完整语义：选择栈非空时验证"账本 = 快照 - 栈内剩余条目的分配明细之和"，
   栈为空（或传 None）时退化为 VBA 的精确相等判定。
3. VBA 用 Err.Raise + vbObjectError 偏移错误号传播 E99；Python 用 E99Error
   异常沿调用链自然上传，由最外层（runner）统一捕获。
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterable, NoReturn

from .ledger import InventoryLedger, InventorySnapshot
from .models import InventoryKey

if TYPE_CHECKING:
    from .strategies import AllocationAttempt

# E99 断言上下文标识（与 VBA RaiseE99 的 context 参数取值一致）
CONTEXT_CONSERVATION = "AssertConservation"
CONTEXT_UNDO_CONSISTENCY = "AssertUndoConsistency"
CONTEXT_UNDO_AFTER_E10 = "AssertUndoConsistency after E10 rollback"


class E99Error(Exception):
    """E99 工程异常（对应 VBA Err.Raise E99_ERROR_NUMBER）。

    属性字段与 VBA RaiseE99 的参数一一对应，消息格式与 VBA 完全一致：
    ``[E99] 库存守恒异常：物流单号=XXX SKU=XXX 期望=N 实际=N 上下文=XXX``
    """

    def __init__(
        self, ship_no: str, sku: str, expected: int, actual: int, context: str
    ) -> None:
        self.ship_no = ship_no
        self.sku = sku
        self.expected = expected
        self.actual = actual
        self.context = context
        super().__init__(
            f"[E99] 库存守恒异常：物流单号={ship_no}"
            f" SKU={sku} 期望={expected} 实际={actual} 上下文={context}"
        )


# -----------------------------------------------------------------------------
# 公开函数
# -----------------------------------------------------------------------------


def assert_conservation(
    snapshot: InventorySnapshot | None,
    ledger: InventoryLedger | None,
    details: Iterable,
    context: str = CONTEXT_CONSERVATION,
) -> bool:
    """验证"分配前快照总量 == 账本当前剩余 + 本次已分配"的守恒等式（§6.5.1）。

    对应 VBA AssertConservation + 失败时的 RaiseE99：
    - snapshot：该 SKU 组分配开始前 take_snapshot 拍摄的快照（只含本组范围）；
    - ledger：分配后的当前账本（只统计快照范围内的键，缺失键按 0 计，
      对应 VBA 的 ledger.Exists 检查）；
    - details：本组全部分配明细（任何带 alloc_qty 属性的对象序列）。

    守恒成立返回 True；被破坏（或 snapshot/ledger 为 None 无法核验）时抛 E99Error。
    """
    if snapshot is None or ledger is None:
        # 无法核验不能视为通过（对应 VBA 返回 False，上层统一升级为 E99）
        raise_e99("", "", 0, 0, context)

    total_before = sum(
        snapshot.get_snapshot_current_qty(key) for key in snapshot.get_ledger_keys()
    )
    total_after = sum(
        ledger.get_current_qty(key) for key in snapshot.get_ledger_keys()
    )
    total_allocated = sum(d.alloc_qty for d in details)

    if total_before != total_after + total_allocated:
        ship_no, sku = _snapshot_scope(snapshot)
        raise_e99(ship_no, sku, total_before, total_after + total_allocated, context)
    return True


def assert_undo_consistency(
    snapshot: InventorySnapshot | None,
    choice_stack: Iterable["AllocationAttempt | None"] | None,
    ledger: InventoryLedger | None,
    context: str = CONTEXT_UNDO_CONSISTENCY,
) -> bool:
    """验证撤销操作后账本状态与"快照 - 选择栈剩余分配"一致（§6.5.2）。

    对应 VBA AssertUndoConsistency，并实现其 choiceStack 预留参数的完整语义：
    - choice_stack 为空或 None：账本每个五元组的当前数量必须与快照精确相等
      （即 VBA 源码实际实现的判定，用于 E10 全量回滚后的检验）；
    - choice_stack 非空：账本 = 快照 - 栈内剩余条目（未撤销行）的分配明细之和，
      用于每次回溯撤销后的实时检验。

    一致返回 True；不一致（或 snapshot/ledger 为 None）时抛 E99Error。
    """
    if snapshot is None or ledger is None:
        raise_e99("", "", 0, 0, context)

    ship_no, sku = _snapshot_scope(snapshot)

    # 汇总选择栈剩余条目按五元组的已分配数量
    committed: dict[InventoryKey, int] = {}
    for entry in choice_stack or ():
        if entry is None:
            continue
        for d in entry.details:
            key = InventoryKey(ship_no, sku, d.qc, d.lot_no, d.expiry)
            committed[key] = committed.get(key, 0) + d.alloc_qty

    # 逐键比对：快照数量 - 剩余已分配 = 账本当前数量
    snapshot_keys = snapshot.get_ledger_keys()
    ledger_keys = set(ledger.get_ledger_keys())
    consistent = True
    for key in snapshot_keys:
        if key not in ledger_keys:
            # 账本中找不到该键本身就是异常（对应 VBA ledger.Exists 检查）
            consistent = False
            break
        expected_qty = snapshot.get_snapshot_current_qty(key) - committed.pop(key, 0)
        if ledger.get_current_qty(key) != expected_qty:
            consistent = False
            break
    # 栈内剩余分配引用了快照范围之外的五元组，同样是异常
    if consistent and any(qty != 0 for qty in committed.values()):
        consistent = False

    if not consistent:
        expected_total = sum(
            snapshot.get_snapshot_current_qty(key) for key in snapshot_keys
        ) - sum(
            d.alloc_qty
            for entry in choice_stack or ()
            if entry is not None
            for d in entry.details
        )
        actual_total = sum(ledger.get_current_qty(key) for key in snapshot_keys)
        raise_e99(ship_no, sku, expected_total, actual_total, context)
    return True


def raise_e99(
    ship_no: str, sku: str, expected: int, actual: int, context: str
) -> NoReturn:
    """触发 E99 工程异常（对应 VBA RaiseE99 的 Err.Raise，调用方用 try/except 捕获）。"""
    raise E99Error(ship_no, sku, expected, actual, context)


# -----------------------------------------------------------------------------
# 私有辅助函数
# -----------------------------------------------------------------------------


def _snapshot_scope(snapshot: InventorySnapshot) -> tuple[str, str]:
    """从快照键中提取（物流单号, SKU）。快照按构造只覆盖单一组；空快照返回空串。"""
    keys = snapshot.get_ledger_keys()
    if not keys:
        return "", ""
    return keys[0].shipment_no, keys[0].sku
