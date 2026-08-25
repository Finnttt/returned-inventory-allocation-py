"""E2E 验收测试：需求 §6.5.3 强制最小测试集 T1~T11+。

全部走 run_full_allocation 端到端（内存工作簿 → 读取/标准化/校验/账本/
排序预检/回溯分配/状态判定/输出写入 全链路），逐字段比对汇总表、
成功分配明细表、数据异常明细表、调试日志表、运行历史记录表的关键输出。

输入与期望优先取自根目录 TC-*.md 冻结数据（TC-01/02/03/11/21/22/24/36/20/19/
32/46 等），与 VBA 版期望保持一致，供阶段 11 对拍复用。

注意（与 TC 文档的差异说明）：TC-19/20/36 文档与冻结 DataSet 的"原因"列带
补充说明括号（如 "E08 - 同物流单号+SKU数量不一致（退单合计=8，库存合计=5）"），
而 VBA modStatus.bas ST_GetStandardReasonText 实际输出的是固定标准文案
（不含括号），Python 移植与 VBA 逐字一致。本文件断言以 VBA 实际行为为准。
"""

from __future__ import annotations

from returned_inventory.excel_output import RUN_HISTORY_HEADERS
from returned_inventory.models import (
    DEBUG_LEVEL_SIMPLE,
    NA_PLACEHOLDER,
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    STATUS_UNALLOCATED,
)
from returned_inventory.runner import run_full_allocation

from .wb_factory import build_test_workbook

# TC 冻结数据使用的物流单号（与根目录 TC-*.md 一致）
SF00 = "SF3190000000000"  # TC-01/30
SF01 = "SF3190000000001"  # TC-02
SF02 = "SF3190000000002"  # TC-03
SF16 = "SF3190000000016"  # TC-21/33/34
SF17 = "SF3190000000017"  # TC-11
SF27 = "SF3190000000027"  # TC-22/23/25a/26/27
SF28 = "SF3190000000028"  # TC-24/25b
SF32 = "SF3190000000032"  # TC-32
SF36 = "SF3190000000036"  # TC-36
SF46 = "SF3190000000046"  # TC-46
SF59 = "SF3190000000059"  # TC-19/38
SF60 = "SF3190000000060"  # TC-20


# -----------------------------------------------------------------------------
# 读取辅助（与 test_runner.py 同款口径；None 归一为 "" 便于逐字段比对）
# -----------------------------------------------------------------------------


def _data_rows(ws):
    return [
        ["" if cell.value is None else cell.value for cell in row]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row)
    ]


def _history_rows(wb):
    return _data_rows(wb["运行历史记录表"])


def _run(order_rows, inventory_rows, config_rows=None):
    """构造内存工作簿并执行完整分配，返回 (wb, stats)。"""
    wb = build_test_workbook(
        order_rows=order_rows, inventory_rows=inventory_rows, config_rows=config_rows
    )
    stats = run_full_allocation(wb)
    return wb, stats


# -----------------------------------------------------------------------------
# T1（TC-01/30）：策略一精确匹配；调试日志关闭；单行库存恰好=需求（T11+ 边界）
# -----------------------------------------------------------------------------


def test_t1_tc01_strategy1_exact_match():
    """TC-01：qty = D = 2 策略一命中；行/退单号状态=批量导入；TC-30 调试日志不生成。"""
    wb, stats = _run(
        order_rows=[[SF00, "TK00000001", "H000000001", "00001", 2]],
        inventory_rows=[[SF00, "H000000001", "ZP", "LA01", "2029/01/01", 2]],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [[SF00, "TK00000001", STATUS_BATCH_IMPORT, ""]]
    assert _data_rows(wb["成功分配明细表"]) == [[
        SF00, "TK00000001", "H000000001", "00001", 2, "ZP", "LA01",
        "2029/01/01", 2, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT,
    ]]
    assert _data_rows(wb["数据异常明细表"]) == []
    assert _data_rows(wb["调试日志"]) == []  # TC-30：级别=关闭，不写数据行

    assert stats.total_backtrack_count == 0
    history = _history_rows(wb)
    assert len(history) == 1
    assert history[0][2] == "Full Run"
    assert history[0][3] == 1 and history[0][4] == 1 and history[0][5] == 1
    assert history[0][9] == 0    # 校验失败物流单号数
    assert history[0][10] == 1   # 分配成功
    assert history[0][11] == 0   # 分配失败
    assert history[0][12] == ""  # 错误码分布
    assert history[0][13] == 0 and history[0][14] == 0  # 总回溯 / 最大单组回溯


# -----------------------------------------------------------------------------
# T2（TC-02）：策略二最接近匹配，单行扣减后剩余库存保留
# -----------------------------------------------------------------------------


def test_t2_tc02_strategy2_closest_match_keeps_remainder():
    """TC-02：行 00001 策略二扣 2（3→1），行 00002 用剩余 1 策略一命中。"""
    wb, stats = _run(
        order_rows=[
            [SF01, "TK00000011", "H000000001", "00001", 2],
            [SF01, "TK00000011", "H000000001", "00002", 1],
        ],
        inventory_rows=[[SF01, "H000000001", "ZP", "LA01", "2029/01/01", 3]],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [[SF01, "TK00000011", STATUS_BATCH_IMPORT, ""]]
    # 行 00002 能分到剩余 1，即证明策略二只扣需求量、未清零整行库存
    assert _data_rows(wb["成功分配明细表"]) == [
        [SF01, "TK00000011", "H000000001", "00001", 2, "ZP", "LA01",
         "2029/01/01", 2, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF01, "TK00000011", "H000000001", "00002", 1, "ZP", "LA01",
         "2029/01/01", 1, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
    ]
    assert stats.total_backtrack_count == 0


# -----------------------------------------------------------------------------
# T3（TC-03）：策略三跨批号/效期拼凑，行状态=手工操作，多条明细
# -----------------------------------------------------------------------------


def test_t3_tc03_strategy3_multi_lot_manual():
    """TC-03：策略三先选效期更晚的 LA02，再选 LA01；两条明细、手工操作。"""
    wb, stats = _run(
        order_rows=[[SF02, "TK00000021", "H000000001", "00001", 2]],
        inventory_rows=[
            [SF02, "H000000001", "ZP", "LA01", "2029/01/01", 1],
            [SF02, "H000000001", "ZP", "LA02", "2030/01/01", 1],
        ],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [[SF02, "TK00000021", STATUS_MANUAL, ""]]
    # 明细行顺序与分配步骤一致：LA02（效期更晚）先、LA01 后；退单数量恒为原始需求 2
    assert _data_rows(wb["成功分配明细表"]) == [
        [SF02, "TK00000021", "H000000001", "00001", 2, "ZP", "LA02",
         "2030/01/01", 1, STATUS_MANUAL, STATUS_MANUAL],
        [SF02, "TK00000021", "H000000001", "00001", 2, "ZP", "LA01",
         "2029/01/01", 1, STATUS_MANUAL, STATUS_MANUAL],
    ]
    assert stats.total_backtrack_count == 0


# -----------------------------------------------------------------------------
# T4（TC-11）：多 QC 竞争 + 静态排序四级规则全触发，明细按处理序输出
# -----------------------------------------------------------------------------


def test_t4_tc11_static_sort_four_levels_e2e():
    """TC-11：可用QC数升序→Qty降序→退单号升序→行号升序；明细顺序=处理序。"""
    wb, stats = _run(
        order_rows=[
            [SF17, "TK00000111", "H000000001", "00001", 10],
            [SF17, "TK00000111", "H000000001", "00002", 8],
            [SF17, "TK00000111", "H000000001", "00003", 4],
            [SF17, "TK00000111", "H000000001", "00004", 5],
            [SF17, "TK00000222", "H000000001", "00001", 5],
            [SF17, "TK00000222", "H000000001", "00002", 5],
        ],
        inventory_rows=[
            [SF17, "H000000001", "ZP", "LA01", "2029/01/01", 12],
            [SF17, "H000000001", "QC", "LA01", "2029/01/01", 20],
            [SF17, "H000000001", "NG", "LA01", "2029/01/01", 5],
        ],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF17, "TK00000111", STATUS_BATCH_IMPORT, ""],
        [SF17, "TK00000222", STATUS_BATCH_IMPORT, ""],
    ]
    # 逐行逐字段比对（TC-11 §5.2）：处理序 00001→00002→00003→00004→222/00001→222/00002
    assert _data_rows(wb["成功分配明细表"]) == [
        [SF17, "TK00000111", "H000000001", "00001", 10, "QC", "LA01", "2029/01/01", 10, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF17, "TK00000111", "H000000001", "00002", 8, "ZP", "LA01", "2029/01/01", 8, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF17, "TK00000111", "H000000001", "00003", 4, "ZP", "LA01", "2029/01/01", 4, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF17, "TK00000111", "H000000001", "00004", 5, "NG", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF17, "TK00000222", "H000000001", "00001", 5, "QC", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF17, "TK00000222", "H000000001", "00002", 5, "QC", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
    ]
    assert stats.total_backtrack_count == 0


# -----------------------------------------------------------------------------
# T5（TC-21/33/34）：回溯触发并成功；五元组汇总；多 SKU 独立分配
# -----------------------------------------------------------------------------


def test_t5_tc21_backtrack_success_five_tuple_merge():
    """TC-21：H1 初始 ZP 路径失败 → 回溯 4 步改选 QC 成功；H2/H3 直接成功。"""
    wb, stats = _run(
        order_rows=[
            [SF16, "TK10000161", "H000000001", "00001", 12],
            [SF16, "TK10000161", "H000000001", "00002", 5],
            [SF16, "TK10000161", "H000000001", "00003", 5],
            [SF16, "TK10000161", "H000000001", "00004", 5],
            [SF16, "TK10000161", "H000000001", "00005", 5],
            [SF16, "TK10000161", "H000000001", "00006", 5],
            [SF16, "TK10000161", "H000000001", "00007", 5],
            [SF16, "TK10000161", "H000000002", "00008", 3],
            [SF16, "TK10000161", "H000000003", "00009", 1],
            [SF16, "TK10000162", "H000000002", "00001", 2],
        ],
        inventory_rows=[
            [SF16, "H000000001", "ZP", "LA01", "2029/01/01", 8],
            [SF16, "H000000001", "ZP", "LA01", "2029/01/01", 12],  # TC-33：五元组合并 8+12=20
            [SF16, "H000000001", "QC", "LA01", "2029/01/01", 12],
            [SF16, "H000000001", "QC", "LA01", "2029/01/01", 5],
            [SF16, "H000000001", "QC", "LA01", "2029/01/01", 5],   # TC-33：合并 12+5+5=22
            [SF16, "H000000002", "NG", "LB01", "2029/01/01", 5],
            [SF16, "H000000003", "NG", "LB01", "2029/01/01", 1],
        ],
        config_rows=[["调试日志级别", DEBUG_LEVEL_SIMPLE, ""]],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF16, "TK10000161", STATUS_BATCH_IMPORT, ""],
        [SF16, "TK10000162", STATUS_BATCH_IMPORT, ""],
    ]
    # TC-21 §四：回溯后 00001~00003 走 QC、00004~00007 走 ZP，终态逐字段比对
    assert _data_rows(wb["成功分配明细表"]) == [
        [SF16, "TK10000161", "H000000001", "00001", 12, "QC", "LA01", "2029/01/01", 12, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000001", "00002", 5, "QC", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000001", "00003", 5, "QC", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000001", "00004", 5, "ZP", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000001", "00005", 5, "ZP", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000001", "00006", 5, "ZP", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000001", "00007", 5, "ZP", "LA01", "2029/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000002", "00008", 3, "NG", "LB01", "2029/01/01", 3, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000162", "H000000002", "00001", 2, "NG", "LB01", "2029/01/01", 2, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF16, "TK10000161", "H000000003", "00009", 1, "NG", "LB01", "2029/01/01", 1, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
    ]

    # 回溯计数（实现口径：单行撤销-重试每步计 1 次，H1 组 = 4）
    assert stats.total_backtrack_count == 4
    history = _history_rows(wb)
    assert history[0][13] == 4 and history[0][14] == 4
    assert history[0][10] == 1 and history[0][11] == 0

    # 简版调试日志：每个退单行 1 条最终结果记录，H1 组记录实际回溯次数 4
    debug = _data_rows(wb["调试日志"])
    assert len(debug) == 10
    assert {row[15] for row in debug if row[1] == "H000000001"} == {4}


# -----------------------------------------------------------------------------
# T6（TC-24/25b）：回溯超限 E10 + 跨 SKU 短路连带回滚 + 整单回滚
# -----------------------------------------------------------------------------


def test_t6_tc24_backtrack_limit_e10_cross_sku_cascade():
    """TC-24：max=10 时第 11 次回溯触发 E10；H2 从未执行，连带回滚。"""
    wb, stats = _run(
        order_rows=[
            [SF28, "TK10000281", "H000000001", "00001", 6],
            [SF28, "TK10000282", "H000000001", "00001", 6],
            [SF28, "TK10000282", "H000000001", "00002", 6],
            [SF28, "TK10000282", "H000000001", "00003", 4],
            [SF28, "TK10000282", "H000000001", "00004", 4],
            [SF28, "TK10000282", "H000000001", "00005", 4],
            [SF28, "TK10000282", "H000000001", "00006", 4],
            [SF28, "TK10000282", "H000000001", "00007", 4],
            [SF28, "TK10000282", "H000000001", "00008", 4],
            [SF28, "TK10000282", "H000000002", "00009", 6],
            [SF28, "TK10000282", "H000000002", "00010", 3],
        ],
        inventory_rows=[
            [SF28, "H000000001", "ZP", "LA01", "2029/01/01", 23],
            [SF28, "H000000001", "QC", "LA01", "2029/01/01", 13],
            [SF28, "H000000001", "NG", "LA01", "2029/01/01", 6],
            [SF28, "H000000002", "ZP", "LA01", "2029/01/01", 9],
        ],
        config_rows=[
            ["最大回溯次数", 10, ""],
            ["调试日志级别", DEBUG_LEVEL_SIMPLE, ""],
        ],
    )

    # 两退单号均含 H1 直接失败行 → 原因均为直接原因格式（TC-42 直接分支）
    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF28, "TK10000281", STATUS_UNALLOCATED, "E10 - 回溯超限"],
        [SF28, "TK10000282", STATUS_UNALLOCATED, "E10 - 回溯超限"],
    ]
    assert _data_rows(wb["成功分配明细表"]) == []  # 整单回滚（TC-26）

    assert stats.alloc_success_count == 0 and stats.alloc_fail_count == 1
    history = _history_rows(wb)
    # TC-24 文档 5.5 填 "E10:1"；VBA 口径（modRunner RN_BuildErrorCodeDistribution）
    # 对 Group_g_ErrorCode 非空即计数，连带回滚组同样计入，故实际为 "E10:1; 连带回滚:1"
    assert history[0][12] == "E10:1; 连带回滚:1"
    assert history[0][13] == 11 and history[0][14] == 11  # 第 11 次回溯超限

    # 简版调试日志：H1 九行 + H2 两行连带回滚（TC-25b 短路总原则）
    # 注：连带回滚行"错误码"列填占位码"连带回滚"（models.ERROR_CASCADE_ROLLBACK，
    # 与 VBA BT_AttachDebugEvents 口径一致），TC-24 文档调试日志表中的 E10/否
    # 为文档旧值，以实现为准。
    debug = _data_rows(wb["调试日志"])
    assert len(debug) == 11
    h2_rows = [row for row in debug if row[1] == "H000000002"]
    assert len(h2_rows) == 2
    assert all(row[18] == "连带回滚—跨SKU短路" for row in h2_rows)
    assert all(row[17] == "连带回滚" and row[15] == 0 for row in h2_rows)  # 从未执行：无回溯


# -----------------------------------------------------------------------------
# T6 补充（TC-22/23/25a/26/27）：回溯路径穷尽 E09（max=200 遍历全部 15 条路径）
# -----------------------------------------------------------------------------


def test_t6b_tc22_e09_all_paths_exhausted():
    """TC-22：15 条路径全失败（实现计步 59）→ E09 → 整单回滚，明细表为空。"""
    wb, stats = _run(
        order_rows=[
            [SF27, "TK10000271", "H000000001", "00001", 6],
            [SF27, "TK10000271", "H000000001", "00002", 6],
            [SF27, "TK10000271", "H000000001", "00003", 6],
            [SF27, "TK10000271", "H000000001", "00004", 4],
            [SF27, "TK10000271", "H000000001", "00005", 4],
            [SF27, "TK10000271", "H000000001", "00006", 4],
            [SF27, "TK10000271", "H000000001", "00007", 4],
            [SF27, "TK10000271", "H000000001", "00008", 4],
            [SF27, "TK10000271", "H000000001", "00009", 4],
        ],
        inventory_rows=[
            [SF27, "H000000001", "ZP", "LA01", "2029/01/01", 23],
            [SF27, "H000000001", "QC", "LA01", "2029/01/01", 13],
            [SF27, "H000000001", "NG", "LA01", "2029/01/01", 6],
        ],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF27, "TK10000271", STATUS_UNALLOCATED, "E09 - 分配路径穷尽"],
    ]
    assert _data_rows(wb["成功分配明细表"]) == []  # TC-26：回滚后无成功明细

    assert stats.alloc_fail_count == 1
    history = _history_rows(wb)
    assert history[0][12] == "E09:1"
    assert history[0][13] == 59 and history[0][14] == 59  # 单行撤销-重试计步


# -----------------------------------------------------------------------------
# T7（TC-36）：E11 QC 库存碎片（0 < T < groupMinQty），校验阶段拦截
# -----------------------------------------------------------------------------


def test_t7_tc36_e11_fragment_inventory():
    """TC-36：ZP=1/QC=1 均为碎片（groupMinQty=2）→ E11，不进入分配阶段。"""
    wb, stats = _run(
        order_rows=[[SF36, "TK00000036", "H000000001", "00001", 2]],
        inventory_rows=[
            [SF36, "H000000001", "ZP", "LA01", "2029/01/01", 1],
            [SF36, "H000000001", "QC", "LA01", "2029/01/01", 1],
        ],
    )

    # 原因以 VBA ST_GetStandardReasonText 标准文案为准（TC-36 文档括号内容为注释性补充）
    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF36, "TK00000036", STATUS_UNALLOCATED, "E11 - QC库存碎片无法分配"],
    ]
    assert _data_rows(wb["成功分配明细表"]) == []
    assert _data_rows(wb["数据异常明细表"]) == []  # E11 仅进汇总表（§5.4）
    assert _data_rows(wb["调试日志"]) == []        # 分配未启动，无日志

    assert stats.validation_fail_count == 1
    assert stats.alloc_success_count == 0 and stats.alloc_fail_count == 0
    history = _history_rows(wb)
    assert history[0][12] == "E11:1"  # 同码同单号去重（ZP/QC 两条碎片记录计 1）
    assert history[0][13] == 0


# -----------------------------------------------------------------------------
# T8（TC-20）：E08 同物流单号+SKU 数量不一致
# -----------------------------------------------------------------------------


def test_t8_tc20_e08_qty_mismatch():
    """TC-20：退单合计 8 ≠ 库存合计 5 → E08；两个退单号均无法分配。"""
    wb, stats = _run(
        order_rows=[
            [SF60, "TK10000600", "H000000060", "00001", 5],
            [SF60, "TK10000601", "H000000060", "00001", 3],
        ],
        inventory_rows=[[SF60, "H000000060", "ZP", "LA01", "2029/06/15", 5]],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF60, "TK10000600", STATUS_UNALLOCATED, "E08 - 同物流单号+SKU数量不一致"],
        [SF60, "TK10000601", STATUS_UNALLOCATED, "E08 - 同物流单号+SKU数量不一致"],
    ]
    assert _data_rows(wb["成功分配明细表"]) == []
    assert _data_rows(wb["数据异常明细表"]) == []  # E08 跨表汇总级，不进异常明细

    assert stats.validation_fail_count == 1
    history = _history_rows(wb)
    assert history[0][12] == "E08:1"
    assert history[0][10] == 0 and history[0][11] == 0


# -----------------------------------------------------------------------------
# T9（TC-25a/26/27/42 连带回滚分支）：部分 SKU 成功、部分失败 → 整单回滚
# -----------------------------------------------------------------------------


def test_t9_partial_sku_success_full_rollback_cascade_reason():
    """整单回滚：H1（WMS-A）本可成功，H2（WMS-B）E09 失败 → 全单撤回。

    原因格式区分（R064/TC-42）：含直接失败 SKU 的 WMS-B 为 "E09 - 分配路径穷尽"；
    仅被连累的 WMS-A 为 "整单回滚（触发原因：E09）"。
    H2 数据采用需求 §4.2.3 预检测B 示例（需求 1/2/5/6，库存 ZP-7/QC-4/NG-3）。
    """
    ship = "SF3190000009999"
    wb, stats = _run(
        order_rows=[
            [ship, "TK10009991", "H000000091", "00001", 2],  # H1 先出现，先分配且本可成功
            [ship, "TK10009992", "H000000092", "00001", 1],
            [ship, "TK10009992", "H000000092", "00002", 2],
            [ship, "TK10009992", "H000000092", "00003", 5],
            [ship, "TK10009992", "H000000092", "00004", 6],
        ],
        inventory_rows=[
            [ship, "H000000091", "ZP", "LA01", "2029/01/01", 2],
            [ship, "H000000092", "ZP", "LA01", "2029/01/01", 7],
            [ship, "H000000092", "QC", "LB01", "2029/01/01", 4],
            [ship, "H000000092", "NG", "LC01", "2029/01/01", 3],
        ],
    )

    summary = {row[1]: row for row in _data_rows(wb["分配状态汇总表"])}
    assert summary["TK10009992"][2] == STATUS_UNALLOCATED
    assert summary["TK10009992"][3] == "E09 - 分配路径穷尽"          # 直接原因
    assert summary["TK10009991"][2] == STATUS_UNALLOCATED
    assert summary["TK10009991"][3] == "整单回滚（触发原因：E09）"    # 连带回滚（TC-42）
    assert _data_rows(wb["成功分配明细表"]) == []  # TC-26：部分成功的 H1 也被撤回

    assert stats.alloc_success_count == 0 and stats.alloc_fail_count == 1
    history = _history_rows(wb)
    assert history[0][12] == "E09:1"


# -----------------------------------------------------------------------------
# T10（TC-19/38）：E07 孤立物流单号，汇总表 WMS 退单号填 [N/A]
# -----------------------------------------------------------------------------


def test_t10_tc19_tc38_e07_orphan_shipment_na_placeholder():
    """TC-19/38：退单表空、库存表有 SF0059 → E07；汇总行 WMS=[N/A]；异常明细 1 行。"""
    wb, stats = _run(
        order_rows=[],
        inventory_rows=[[SF59, "H000000059", "ZP", "LA01", "2029/06/15", 5]],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [
        [SF59, NA_PLACEHOLDER, STATUS_UNALLOCATED, "E07 - 物流单号仅存在于质检库存表"],
    ]
    assert _data_rows(wb["成功分配明细表"]) == []
    assert _data_rows(wb["数据异常明细表"]) == [[
        "质检库存表", 2, SF59, NA_PLACEHOLDER, "H000000059",
        "物流单号", SF59, "E07", "物流单号仅存在于质检库存表",
    ]]

    assert stats.validation_fail_count == 1
    assert stats.input_return_rows == 0 and stats.input_inventory_rows == 1
    history = _history_rows(wb)
    assert history[0][12] == "E07:1"


# -----------------------------------------------------------------------------
# T11+ 边界：效期哨兵值（TC-32）
# -----------------------------------------------------------------------------


def test_t11_tc32_no_expiry_sentinel_accepted():
    """TC-32：哨兵效期 2099/01/01 不触发 E05，参与分配并原样输出。"""
    wb, stats = _run(
        order_rows=[
            [SF32, "TK10000320", "H000000032", "00001", 5],
            [SF32, "TK10000320", "H000000032", "00002", 5],
        ],
        inventory_rows=[[SF32, "H000000032", "ZP", "LA01", "2099/01/01", 10]],
    )

    assert _data_rows(wb["分配状态汇总表"]) == [[SF32, "TK10000320", STATUS_BATCH_IMPORT, ""]]
    assert _data_rows(wb["成功分配明细表"]) == [
        [SF32, "TK10000320", "H000000032", "00001", 5, "ZP", "LA01", "2099/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
        [SF32, "TK10000320", "H000000032", "00002", 5, "ZP", "LA01", "2099/01/01", 5, STATUS_BATCH_IMPORT, STATUS_BATCH_IMPORT],
    ]
    assert stats.total_backtrack_count == 0


# -----------------------------------------------------------------------------
# T11+ 边界：批号大小写（TC-46）
# -----------------------------------------------------------------------------


def test_t11_tc46_lot_case_insensitive_merge():
    """TC-46：批号 a01/A01 标准化合并为 T=10；分配成功且输出统一大写 A01。"""
    wb, stats = _run(
        order_rows=[
            [SF46, "TK10000460", "H000000046", "00001", 5],
            [SF46, "TK10000460", "H000000046", "00002", 5],
        ],
        inventory_rows=[
            [SF46, "H000000046", "ZP", "a01", "2029/06/15", 6],
            [SF46, "H000000046", "ZP", "A01", "2029/06/15", 4],
        ],
    )

    # 若未合并（T=6/T=4 分裂），两行均不可分配 → E09；分配成功即反证 R021 生效
    assert _data_rows(wb["分配状态汇总表"]) == [[SF46, "TK10000460", STATUS_BATCH_IMPORT, ""]]
    detail = _data_rows(wb["成功分配明细表"])
    assert [row[6] for row in detail] == ["A01", "A01"]  # 批号统一大写输出
    assert [row[8] for row in detail] == [5, 5]
    assert stats.alloc_success_count == 1


# -----------------------------------------------------------------------------
# T11+ 边界：批号前导零保留（无独立 TC 文档，§5.2/EO_ApplyTextFormats 口径）
# -----------------------------------------------------------------------------


def test_t11_lot_leading_zeros_preserved_end_to_end():
    """批号 "0007" 全链路保留前导零：不被数值化，明细表原样输出文本 "0007"。"""
    ship = "SF3190000009998"
    wb, stats = _run(
        order_rows=[[ship, "TK10009998", "H000000098", "00001", 3]],
        inventory_rows=[[ship, "H000000098", "ZP", "0007", "2029/01/01", 3]],
    )

    assert stats.alloc_success_count == 1
    detail = _data_rows(wb["成功分配明细表"])
    assert detail[0][6] == "0007"   # 批号列
    assert isinstance(detail[0][6], str)
    assert detail[0][3] == "00001"  # 行号前导零同样保留


# -----------------------------------------------------------------------------
# T11+ 边界：空行（无独立 TC 文档，按 R041/E01 规则推导，见下方断言注释）
# -----------------------------------------------------------------------------


def test_t11_blank_middle_row_flagged_e01_others_unaffected():
    """中间空行触发 E01（字段级异常），但不拖垮同表其他合法物流单号。

    行为口径（与 VBA 对齐，并经 SF0013 冻结预期中 [N/A] 汇总行模式佐证）：
    - 空行五个关键字段均为空 → 逐字段生成 E01，进入数据异常明细表（Excel行号=3）；
    - 空行无法归属任何物流单号 → 汇总表落一行 [N/A]/[N/A]/无法分配/E01；
    - failed_shipments 不含 [N/A]，合法物流单号 SF...997 照常分配成功；
    - 末尾空行超出最后使用行范围，根本不参与读取（异常明细只有行 3）。
    """
    ship = "SF3190000009997"
    wb, stats = _run(
        order_rows=[
            [ship, "TK10009997", "H000000097", "00001", 2],
            [None, None, None, None, None],  # 中间空行
            [ship, "TK10009997", "H000000097", "00002", 3],
            [None, None, None, None, None],  # 末尾空行（最后使用行之外，不读取）
        ],
        inventory_rows=[[ship, "H000000097", "ZP", "LA01", "2029/01/01", 5]],
    )

    anomaly = _data_rows(wb["数据异常明细表"])
    assert len(anomaly) == 5  # 空行 5 个关键字段各一条 E01
    assert all(row[1] == 3 and row[7] == "E01" for row in anomaly)

    summary = {row[0]: row for row in _data_rows(wb["分配状态汇总表"])}
    assert summary[NA_PLACEHOLDER][1] == NA_PLACEHOLDER
    assert summary[NA_PLACEHOLDER][2] == STATUS_UNALLOCATED
    assert summary[NA_PLACEHOLDER][3] == "E01 - 关键字段为空或格式异常"
    assert summary[ship][2] == STATUS_BATCH_IMPORT  # 合法单不受空行牵连

    assert stats.validation_fail_count == 0  # [N/A] 不计入失败物流单号数
    assert stats.alloc_success_count == 1
    assert len(_data_rows(wb["成功分配明细表"])) == 2
