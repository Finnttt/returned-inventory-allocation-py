"""M06 库存账本（对应 VBA modInventoryLedger.bas）。

职责：按五元组（物流单号+SKU+QC+批号+效期）汇总质检库存，
提供统一的库存状态管理接口，是整个系统中库存状态的唯一持有者。

内部数据结构：
- InventoryLedger 内部为 dict[InventoryKey, list[int]]，
  value 为 [原始数量, 当前可用数量]，对应 VBA 的 Array(原始数量, 当前可用数量)。
  Python dict 保持插入顺序，与 VBA Scripting.Dictionary 的遍历顺序语义一致。
- InventorySnapshot 与账本结构相同，但只含指定（物流单号+SKU）范围的深拷贝副本。
- undo_log 为 list，每次 deduct 追加一条 (InventoryKey, 扣减数量) 记录，
  undo 时按 LIFO 顺序还原（对应 VBA Collection + "键&分隔符&数量" 字符串记录）。

与 VBA 的有意偏差：VBA 用 vbNullChar 拼接字符串做键、用字符串前缀做范围匹配；
Python 直接用五元组 InventoryKey 做 dict 键、用字段相等比较做范围过滤，
语义完全等价且不存在分隔符与业务数据冲突的问题。
"""

from __future__ import annotations

from .models import InventoryKey, InventoryRow, NormalizedInventoryLine

# value 数组下标（对应 VBA IDX_ORIG / IDX_CURR）
IDX_ORIG = 0
IDX_CURR = 1

# 撤销日志：每次 deduct 追加一条 (五元组键, 扣减数量)，undo 按 LIFO 还原。
UndoLog = list[tuple[InventoryKey, int]]


# -----------------------------------------------------------------------------
# 核心公开函数
# -----------------------------------------------------------------------------


def build_ledger(inventory: list[NormalizedInventoryLine]) -> InventoryLedger:
    """建立库存账本：将 NormalizedInventoryLine 列表按五元组汇总。

    只有三个合法性标记（qc_valid/expiry_valid/qty_valid）全部为 True 的行
    才计入账本；非法行已由 validate 报错，这里只纳入"可用库存"。
    """
    entries: dict[InventoryKey, list[int]] = {}
    for line in inventory:
        if not (line.qc_valid and line.expiry_valid and line.qty_valid):
            continue
        key = InventoryKey(line.shipment_no, line.sku, line.qc, line.lot_no, line.expiry)
        if key in entries:
            # 相同五元组：累加数量（同一批货可能分多行录入）
            entries[key][IDX_ORIG] += line.qty
            entries[key][IDX_CURR] += line.qty
        else:
            entries[key] = [line.qty, line.qty]
    return InventoryLedger(entries)


def new_undo_log() -> UndoLog:
    """创建新的空撤销日志（对应 VBA NewUndoLog）。"""
    return []


class InventoryLedger:
    """库存账本：五元组 → [原始数量, 当前可用数量] 的有序映射。"""

    def __init__(self, entries: dict[InventoryKey, list[int]]) -> None:
        self._entries = entries

    def query_qc_total(self, ship_no: str, sku: str, qc: str) -> int:
        """查询特定（物流单号, SKU, QC）三元组下的当前可用总量。

        该三元组下可能有多个批号/效期（不同五元组），本函数将它们全部加总。
        M07/M08 用此值判断某 QC 是否有足够库存满足退单需求。
        """
        return sum(
            entry[IDX_CURR]
            for key, entry in self._entries.items()
            if key.shipment_no == ship_no and key.sku == sku and key.qc == qc
        )

    def get_five_tuple_rows(self, ship_no: str, sku: str, qc: str) -> list[InventoryRow]:
        """获取特定（物流单号, SKU, QC）下所有五元组行（含当前可用数量）。

        M07/M08 用此结果构建候选批号/效期列表，每行对应一个可供选择的库存格。
        """
        return [
            InventoryRow(
                shipment_no=key.shipment_no,
                sku=key.sku,
                qc=key.qc,
                lot_no=key.lot_no,
                expiry=key.expiry,
                original_qty=entry[IDX_ORIG],
                current_qty=entry[IDX_CURR],
            )
            for key, entry in self._entries.items()
            if key.shipment_no == ship_no and key.sku == sku and key.qc == qc
        ]

    def deduct(self, key: InventoryKey, qty: int, undo_log: UndoLog) -> bool:
        """从账本中扣减指定五元组的库存数量。

        成功：返回 True，current_qty 减少，undo_log 追加一条还原记录。
        失败（库存不足或 key 不存在）：返回 False，账本和 undo_log 均不变。
        这是库存守恒的关键入口——所有分配操作都必须通过此方法修改账本。
        """
        entry = self._entries.get(key)
        if entry is None:
            return False
        if entry[IDX_CURR] < qty:
            # 可用数量不足，拒绝扣减，账本保持原状（防止超发）
            return False
        entry[IDX_CURR] -= qty
        undo_log.append((key, qty))
        return True

    def undo(self, undo_log: UndoLog) -> None:
        """撤销 undo_log 中记录的所有扣减操作。

        按 LIFO（后进先出）顺序还原，确保与扣减顺序完全对称。
        撤销完成后清空 undo_log，防止被重复调用导致数量异常增加。
        """
        while undo_log:
            key, qty = undo_log.pop()
            entry = self._entries.get(key)
            if entry is not None:
                entry[IDX_CURR] += qty

    def take_snapshot(self, ship_no: str, sku: str) -> InventorySnapshot:
        """快照：将指定（物流单号, SKU）范围内所有五元组的当前状态复制为独立副本。

        快照是深拷贝——账本后续的任何 deduct/undo 都不会影响快照内容。
        M10 守卫在分配一个 SKU 组之前先取快照，分配完成后用快照验证守恒等式。
        """
        return InventorySnapshot(
            {
                key: entry.copy()
                for key, entry in self._entries.items()
                if key.shipment_no == ship_no and key.sku == sku
            }
        )

    # -------------------------------------------------------------------------
    # 辅助公开方法（对应 VBA 第二节，供 M10 守卫和测试调用）
    # -------------------------------------------------------------------------

    def get_current_qty(self, key: InventoryKey) -> int:
        """读取账本中指定五元组的当前可用数量；key 不存在时返回 0。"""
        entry = self._entries.get(key)
        return 0 if entry is None else entry[IDX_CURR]

    def get_original_qty(self, key: InventoryKey) -> int:
        """读取账本中指定五元组的原始数量（建账本时写入，此后不变）。"""
        entry = self._entries.get(key)
        return 0 if entry is None else entry[IDX_ORIG]

    def get_ledger_keys(self) -> list[InventoryKey]:
        """获取账本中所有五元组键，供 M10 守卫遍历守恒验证使用。"""
        return list(self._entries.keys())


class InventorySnapshot:
    """库存快照（对应 VBA TakeSnapshot 的返回字典）。

    结构与账本相同，但只含指定（物流单号+SKU）范围的深拷贝副本，
    与账本完全独立。
    """

    def __init__(self, entries: dict[InventoryKey, list[int]]) -> None:
        self._entries = entries

    def get_snapshot_current_qty(self, key: InventoryKey) -> int:
        """读取快照中指定五元组的当前可用数量（即分配前的可用量）。"""
        entry = self._entries.get(key)
        return 0 if entry is None else entry[IDX_CURR]

    def get_ledger_keys(self) -> list[InventoryKey]:
        """获取快照中所有五元组键，供 M10 守卫遍历守恒验证使用。"""
        return list(self._entries.keys())
