"""M09 回溯分配引擎（对应 VBA modBacktracking.bas，需求 §4.2 / §4.2.7 / §6.3.1.3 / §6.5）。

职责：提供两层接口，以物流单号为粒度完成带回溯的分配计算。

- allocate_shipment（公开）：遍历物流单号下的 SKU 组，逐组调用 _allocate_sku_group；
  某组返回 E09/E10 时立即短路（§4.2 短路总原则），剩余 SKU 组生成
  error_code=连带回滚 的 GroupAllocResult。E99 以 E99Error 异常形式沿调用链
  自然上传（对应 VBA Err.Raise 直传 M15，本函数不捕获）。
- _allocate_sku_group（私有）：单 SKU 组的深度优先回溯分配。选择栈 +
  attempted_qcs_by_row + 回溯计数，逐行对齐 VBA AllocateSKUGroup。

回溯语义（§4.2.7，与 VBA 逐行对应，0-based 下标）：
  a. 当前行失败 → 回退到选择栈栈顶行（prev_row = current_row - 1）。VBA 实现为
     "单行撤销-重试"逐级回退：目标行若无剩余可选 QC，会在重试时再次失败并继续
     回退，效果等价于 spec 伪代码的"找最近一个当前可选QC数 > 已尝试QC数的 target 行"。
     回溯计数口径 = 每回退一行计 1 次（见《调试日志19列规格说明.md》2026-07-19 口径统一）。
  b. 撤销 prev_row 的入栈行，用其 undo_log 恢复库存（ledger.undo）。
  c. prev_row 上次用的 QC 记入 attempted_qcs_by_row[prev_row]（去重，保插入序）。
  d. ★清除 prev_row 之后所有行的已尝试QC记录★（漏掉会漏解，§4.2.7 步骤 4）。
  e. 指针回退到 prev_row 重新分配；回溯计数 +1（在撤销前自增并先判超限）；
     超过 cfg.max_backtrack_count → 撤销所有已入栈行后返回 E10。
  f. 第一行就失败（无可回溯的历史行）→ E09（回溯路径穷尽）。

守卫挂载点（§6.5）：组开始前取快照；每次回溯撤销后 assert_undo_consistency；
E10 全量回滚后再次 assert_undo_consistency；组结束后 assert_conservation。

与 VBA 的有意偏差：
1. VBA 的 StaticPlan/attempt/结果均为 Scripting.Dictionary 平铺键；Python 用
   StaticPlan / AllocationAttempt / GroupAllocResult / ShipmentAllocResult
   dataclass 承载，字段一一对应。
2. 行状态判定按"实际使用的(批号+效期)组合数"（1 → 批量导入；≥2 → 手工操作，
   需求 §4.2.4 步骤 3 / §4.4.1"状态由实际结果决定"）；VBA BT_GetLineStatus 按
   策略名判定。在当前五元汇总模型下两者可证明等价（策略三必然 ≥2 个组合）。
3. 每次回溯撤销后执行栈感知的 assert_undo_consistency（§6.5.2 撤销点检验）；
   VBA 仅在 E10 全量回滚后断言（其 AssertUndoConsistency 的 choiceStack 参数
   预留未用）。本挂载能在撤销出错时更早定位，正常流程下不改变任何结果。
4. 已知 VBA 实现与旧版冻结期望日志表的差异，本移植一律以 VBA 代码为准：
   - 最终结果事件的"批号/效期组合数"列固定写 "1"（VBA BT_FillSuccessDebugFields
     硬编码，即使策略三多组合行也是如此）；
   - 失败事件子类型的分支顺序：处理序 >= 首个失败行且非首行的事件一律记
     "连带回滚—同SKU未到达行"（VBA BT_FillFailureDebugFields 分支顺序如此，
     与 TC-22/TC-24 文档中较早的冻结日志表叙述不同）；
   - 成功事件按"行号"首个匹配定位明细（VBA BT_FindDetailIndexByLineNo），
     同组跨退单号行号重复时可能命中首条（潜在 VBA 行为，原样保留）。
"""

from __future__ import annotations

from .guards import (
    CONTEXT_UNDO_AFTER_E10,
    assert_conservation,
    assert_undo_consistency,
)
from .ledger import InventoryLedger
from .models import (
    DEBUG_LEVEL_DETAIL,
    DEBUG_LEVEL_OFF,
    ERR_E09,
    ERR_E10,
    ERROR_CASCADE_ROLLBACK,
    LINE_STATUS_FAILED,
    STATUS_BATCH_IMPORT,
    STATUS_MANUAL,
    AllocationDetail,
    AllocationEvent,
    Config,
    GroupAllocResult,
    GroupStats,
    PrecheckResult,
    ShipmentAllocResult,
)
from .sort_filter import StaticPlan, filter_candidate_pool
from .strategies import AllocationAttempt, try_allocate

# 预检测命中标记（GroupStats.precheck_hit 取值，对应 VBA "预检测A"/"预检测B"）
PRECHECK_HIT_A = "预检测A"
PRECHECK_HIT_B = "预检测B"

# 分配失败子类型（调试日志第 19 列枚举，见《调试日志19列规格说明.md》§3）
FAIL_SUB_PRECHECK_A = "预检测A（初始可用QC=0）"
FAIL_SUB_PRECHECK_B = "预检测B（强制竞争库存不足）"
FAIL_SUB_NO_AVAILABLE_QC = "动态分配无可用QC"
FAIL_SUB_PATH_EXHAUSTED = "回溯路径穷尽"
FAIL_SUB_CASCADE_SAME_SKU = "连带回滚—同SKU未到达行"
FAIL_SUB_CASCADE_CROSS_SKU = "连带回滚—跨SKU短路"

# 过程事件行状态（仅详细模式产生）
PROCESS_LINE_STATUS_ATTEMPT_OK = "过程-尝试成功"
PROCESS_LINE_STATUS_ATTEMPT_FAIL = "过程-尝试失败"
PROCESS_LINE_STATUS_REVOKE = "过程-回溯撤销"


# -----------------------------------------------------------------------------
# 一、公开函数：allocate_shipment
# -----------------------------------------------------------------------------


def allocate_shipment(
    ship_no: str,
    sku_list: list[str],
    plan_map: dict[str, StaticPlan],
    precheck_map: dict[str, PrecheckResult],
    ledger: InventoryLedger,
    cfg: Config,
) -> ShipmentAllocResult:
    """以物流单号为单位，遍历所有 SKU 组完成分配（对应 VBA AllocateShipment）。

    短路逻辑（§4.2 短路总原则）：某 SKU 组返回 E09/E10 时立即停止遍历，
    为后续未处理的 SKU 组生成 error_code=连带回滚 的 GroupAllocResult
    （backtrack_count=0、precheck_hit 空、无明细；调试日志非关闭且该 SKU
    有静态计划时仍生成最终结果事件，子类型=连带回滚—跨SKU短路）。

    E99 不在此捕获：守卫断言失败抛出的 E99Error 沿调用链自然上传，
    对应 VBA 中 Err.Raise 直传 M15 统一处理。

    参数：
    - plan_map：键=SKU，值=build_static_plan 返回的 StaticPlan；
      找不到计划的 SKU 视为 E09（对应 VBA 传空 Dictionary）。
    - precheck_map：键=SKU，值=run_precheck 返回的 PrecheckResult；
      缺失时按 (False, False) 处理（对应 VBA Array(False, False)）。
    """
    result = ShipmentAllocResult(shipment_no=ship_no)
    short_circuit = False

    for sku in sku_list:
        if short_circuit:
            # 前面某组已失败，本组标记"连带回滚"，不再分配
            stats = GroupStats(ship_no, sku, 0, "")
            group = GroupAllocResult(
                shipment_no=ship_no,
                sku=sku,
                success=False,
                error_code=ERROR_CASCADE_ROLLBACK,
                stats=stats,
            )
            plan = plan_map.get(sku)
            if plan is not None:
                group.events = _attach_debug_events(group, plan, cfg, cross_sku_short=True)
            result.group_results.append(group)
            continue

        plan = plan_map.get(sku)
        precheck = precheck_map.get(sku, PrecheckResult())
        group = _allocate_sku_group(plan, precheck, ledger, cfg)
        result.group_results.append(group)

        if group.error_code in (ERR_E09, ERR_E10):
            short_circuit = True

    return result


# -----------------------------------------------------------------------------
# 二、私有函数：_allocate_sku_group
# -----------------------------------------------------------------------------


def _allocate_sku_group(
    plan: StaticPlan | None,
    precheck_result: PrecheckResult,
    ledger: InventoryLedger,
    cfg: Config,
) -> GroupAllocResult:
    """以单个 SKU 组为单位，使用深度优先回溯算法完成分配（对应 VBA AllocateSKUGroup）。

    成功：success=True，details 含全部行明细，stats.backtrack_count 为实际回溯次数。
    失败：error_code = E09（预检测命中/回溯路径穷尽）或 E10（回溯超限），details 为空。
    守卫断言失败：抛 E99Error（对应 VBA RaiseE99），不由本函数捕获。
    """
    # 防御：plan 为空时快速失败（对应 VBA plan Is Nothing / 不含 RowCount / RowCount=0，
    # 此分支 VBA 不产生调试事件）
    if plan is None or plan.row_count == 0:
        ship_no = plan.shipment_no if plan is not None else ""
        sku = plan.sku if plan is not None else ""
        return GroupAllocResult(
            shipment_no=ship_no,
            sku=sku,
            success=False,
            error_code=ERR_E09,
            stats=GroupStats(ship_no, sku, 0, ""),
        )

    ship_no = plan.shipment_no
    sku = plan.sku
    stats = GroupStats(ship_no, sku, 0, "")
    result = GroupAllocResult(
        shipment_no=ship_no, sku=sku, success=False, error_code="", stats=stats
    )

    # --- 预检测结论：命中则整组 E09，跳过分配循环（§4.2.3）---
    if precheck_result.precheck_a_hit:
        result.error_code = ERR_E09
        stats.precheck_hit = PRECHECK_HIT_A
        result.events = _attach_debug_events(result, plan, cfg)
        return result
    if precheck_result.precheck_b_hit:
        result.error_code = ERR_E09
        stats.precheck_hit = PRECHECK_HIT_B
        result.events = _attach_debug_events(result, plan, cfg)
        return result

    # --- 拍摄快照（分配前的库存状态，供守卫断言使用）---
    snapshot = ledger.take_snapshot(ship_no, sku)

    # --- 初始化回溯状态 ---
    row_count = plan.row_count
    backtrack_count = 0
    # choice_stack[r]：第 r 行分配成功的 AllocationAttempt（对应 VBA choiceStack）
    choice_stack: list[AllocationAttempt | None] = [None] * row_count
    # attempted_qcs_by_row[r]：第 r 行已尝试过的 QC（保插入序，对应 VBA 逗号分隔串）
    attempted_qcs_by_row: list[list[str]] = [[] for _ in range(row_count)]
    # 过程事件（仅详细模式产生；简版只保留最终结果行）
    process_events: list[AllocationEvent] = []

    # --- 主分配/回溯循环 ---
    current_row = 0
    while 0 <= current_row < row_count:
        line = plan.lines[current_row]
        tried = attempted_qcs_by_row[current_row]

        # 动态筛选候选池（已尝试的 QC 被排除在外），然后依次尝试三级策略
        pool = filter_candidate_pool(line.line_key, plan, ledger, tried)
        attempt = try_allocate(pool, line.qty, ledger)

        if attempt.success:
            _append_process_attempt_event(
                process_events, cfg, plan, current_row, pool, tried, attempt,
                backtrack_count, PROCESS_LINE_STATUS_ATTEMPT_OK, "", "",
            )
            # 分配成功：推入选择栈，前进到下一行
            choice_stack[current_row] = attempt
            current_row += 1
            continue

        _append_process_attempt_event(
            process_events, cfg, plan, current_row, pool, tried, attempt,
            backtrack_count, PROCESS_LINE_STATUS_ATTEMPT_FAIL,
            ERR_E09, FAIL_SUB_NO_AVAILABLE_QC,
        )

        # 分配失败（候选池为空或三级策略均不满足）
        if current_row == 0:
            # 第一行就失败，没有可回溯的历史行 → 彻底无解，E09
            current_row = -1  # 置为 -1 作为"耗尽"标志，退出循环后检测
            break

        # 回溯：撤销上一行的提交，换一个 QC 重试
        backtrack_count += 1

        if backtrack_count > cfg.max_backtrack_count:
            # 超过回溯上限 → E10；必须先撤销所有已入栈行，保证账本恢复原状
            for ri in range(current_row - 1, -1, -1):
                committed = choice_stack[ri]
                if committed is not None:
                    _append_process_revoke_event(
                        process_events, cfg, plan, ri, committed, backtrack_count
                    )
                    ledger.undo(committed.undo_log)
                    choice_stack[ri] = None

            result.error_code = ERR_E10
            stats.backtrack_count = backtrack_count

            # 守卫检查：E10 回滚后账本应与快照完全一致（栈为空，精确相等判定）
            assert_undo_consistency(snapshot, None, ledger, CONTEXT_UNDO_AFTER_E10)

            result.events = _attach_debug_events_with_process(
                result, plan, cfg, process_events
            )
            return result

        # 撤销上一行（prev_row = current_row - 1）的已提交分配
        prev_row = current_row - 1
        committed = choice_stack[prev_row]
        # 该行上次使用的 QC（策略一/二/三都只用同一个 QC，对应 VBA 取 QC_1）
        used_qc = committed.used_qc

        _append_process_revoke_event(
            process_events, cfg, plan, prev_row, committed, backtrack_count
        )
        ledger.undo(committed.undo_log)
        choice_stack[prev_row] = None

        # 把刚用过的 QC 加入该行的已尝试列表，下次跳过它（去重）
        if used_qc not in attempted_qcs_by_row[prev_row]:
            attempted_qcs_by_row[prev_row].append(used_qc)

        # ★清除 prev_row 之后所有行的已尝试记录★（§4.2.7 步骤 4，漏掉会漏解）
        for ri in range(prev_row + 1, row_count):
            attempted_qcs_by_row[ri] = []

        # 守卫检查（§6.5.2 撤销点检验）：账本 = 快照 - 栈内剩余条目的分配明细
        assert_undo_consistency(snapshot, choice_stack, ledger)

        # 回到 prev_row 重新尝试
        current_row = prev_row

    # --- 循环结束后判断结果 ---
    if current_row >= row_count:
        # 所有行均成功分配，构建 AllocationDetail 列表
        details = _build_details(plan, choice_stack)

        # 守卫断言（§6.5.1 组结束点）：验证库存守恒等式，失败抛 E99Error
        assert_conservation(snapshot, ledger, details)

        result.success = True
        stats.backtrack_count = backtrack_count
        result.details = details
    else:
        # current_row < 0：第一行就失败，彻底无解 → E09
        result.error_code = ERR_E09
        stats.backtrack_count = backtrack_count

    result.events = _attach_debug_events_with_process(result, plan, cfg, process_events)
    return result


# -----------------------------------------------------------------------------
# 三、私有辅助函数：明细构建与行状态
# -----------------------------------------------------------------------------


def _build_details(
    plan: StaticPlan, choice_stack: list[AllocationAttempt | None]
) -> list[AllocationDetail]:
    """从选择栈提取所有行的分配结果，合并为 AllocationDetail 列表
    （对应 VBA BT_BuildDetails）。行状态按实际(批号+效期)组合数判定。"""
    details: list[AllocationDetail] = []
    for r, line in enumerate(plan.lines):
        attempt = choice_stack[r]
        if attempt is None:
            continue
        line_status = _line_status_by_combo_count(attempt)
        for d in attempt.details:
            details.append(
                AllocationDetail(
                    shipment_no=plan.shipment_no,
                    wms_order_no=line.wms_order_no,
                    sku=plan.sku,
                    line_no=line.line_no,
                    order_qty=line.qty,
                    qc=d.qc,
                    lot_no=d.lot_no,
                    expiry=d.expiry,
                    alloc_qty=d.alloc_qty,
                    line_status=line_status,
                    strategy_used=attempt.strategy_used,
                )
            )
    return details


def _line_status_by_combo_count(attempt: AllocationAttempt) -> str:
    """按实际使用的(批号+效期)组合数判定行状态（§4.2.4 步骤 3 / §4.4.1）。

    恰好 1 种 → 批量导入；2 种及以上 → 手工操作。
    对应 VBA BT_GetLineStatus 的策略名判定（两者在当前五元汇总模型下等价，
    见模块头"与 VBA 的有意偏差"第 2 条）。
    """
    combos = {(d.lot_no, d.expiry) for d in attempt.details}
    return STATUS_BATCH_IMPORT if len(combos) <= 1 else STATUS_MANUAL


# -----------------------------------------------------------------------------
# 四、调试日志事件（19 列，M09 → M13，见《调试日志19列规格说明.md》）
# -----------------------------------------------------------------------------


def _attach_debug_events_with_process(
    result: GroupAllocResult,
    plan: StaticPlan,
    cfg: Config,
    process_events: list[AllocationEvent],
) -> list[AllocationEvent]:
    """带过程事件的调试日志组装（对应 VBA BT_AttachDebugEventsWithProcess）。

    详细模式：过程事件在前、最终结果事件在后；简版只保留最终结果事件；
    关闭模式不产生任何事件。
    """
    final_events = _attach_debug_events(result, plan, cfg)
    if cfg.debug_log_level != DEBUG_LEVEL_DETAIL or not process_events:
        return final_events
    return process_events + final_events


def _attach_debug_events(
    result: GroupAllocResult,
    plan: StaticPlan | None,
    cfg: Config,
    cross_sku_short: bool = False,
) -> list[AllocationEvent]:
    """为每个计划行构建最终结果事件（对应 VBA BT_AttachDebugEvents）。

    调试日志关闭、plan 为空或行数为 0 时不产生事件。
    """
    if cfg.debug_log_level == DEBUG_LEVEL_OFF:
        return []
    if plan is None or plan.row_count == 0:
        return []

    ship_no = plan.shipment_no
    sku = plan.sku
    row_count = plan.row_count
    success = result.success
    error_code = result.error_code
    backtrack_count = result.stats.backtrack_count
    precheck_hit = result.stats.precheck_hit

    fail_sub_type = _resolve_fail_sub_type(precheck_hit, error_code, cross_sku_short)
    first_fail_row = _find_first_fail_row(plan, success, error_code, precheck_hit)

    events: list[AllocationEvent] = []
    for r, line in enumerate(plan.lines):
        evt = AllocationEvent(
            shipment_no=ship_no,
            sku=sku,
            wms_order_no=line.wms_order_no,
            line_no=line.line_no,
            demand_d=line.qty,
            process_order=str(r + 1),
            dynamic_next_min_qty=_format_next_min_qty(plan, r),
            candidate_qc_count=str(line.init_qc_count),
            excluded_qc_list="",
            strategy_used="",
            used_qc="",
            qc_before="",
            qc_after="",
            lot_expiry_combo_count="",
            is_backtrack_retry="",
            backtrack_no=backtrack_count,
            line_status="",
            error_code="",
            fail_sub_type="",
            is_final_result=True,
            is_revoked=False,
        )
        if success:
            _fill_success_debug_fields(evt, result, line, backtrack_count)
        else:
            _fill_failure_debug_fields(
                evt, error_code, fail_sub_type, r, first_fail_row,
                backtrack_count, cross_sku_short,
            )
        events.append(evt)
    return events


def _append_process_attempt_event(
    events: list[AllocationEvent],
    cfg: Config,
    plan: StaticPlan | None,
    row_index: int,
    pool: list,
    tried_qcs: list[str],
    attempt: AllocationAttempt | None,
    backtrack_count: int,
    line_status: str,
    error_code: str,
    fail_sub_type: str,
) -> None:
    """记录一次分配尝试过程（对应 VBA BT_AppendProcessAttemptEvent，仅详细模式）。"""
    if cfg.debug_log_level != DEBUG_LEVEL_DETAIL or plan is None:
        return

    evt = _new_common_process_event(plan, row_index, backtrack_count)
    evt.candidate_qc_count = str(len({row.qc for row in pool}))
    evt.excluded_qc_list = ",".join(tried_qcs)
    evt.line_status = line_status
    evt.error_code = error_code
    evt.fail_sub_type = fail_sub_type
    evt.is_final_result = False
    evt.is_revoked = False

    if attempt is None:
        evt.strategy_used = "-"
        evt.used_qc = "-"
        evt.lot_expiry_combo_count = "-"
    else:
        evt.strategy_used = attempt.strategy_used
        evt.used_qc = _attempt_qc_list(attempt)
        evt.lot_expiry_combo_count = str(len(attempt.details))

    events.append(evt)


def _append_process_revoke_event(
    events: list[AllocationEvent],
    cfg: Config,
    plan: StaticPlan | None,
    row_index: int,
    attempt: AllocationAttempt | None,
    backtrack_count: int,
) -> None:
    """记录一次回溯撤销过程（对应 VBA BT_AppendProcessRevokeEvent，仅详细模式）。

    撤销前记录，便于详细日志看到被撤回的选择。
    """
    if cfg.debug_log_level != DEBUG_LEVEL_DETAIL or plan is None or attempt is None:
        return

    evt = _new_common_process_event(plan, row_index, backtrack_count)
    evt.candidate_qc_count = "-"
    evt.excluded_qc_list = "-"
    evt.strategy_used = attempt.strategy_used
    evt.used_qc = _attempt_qc_list(attempt)
    evt.lot_expiry_combo_count = str(len(attempt.details))
    evt.line_status = PROCESS_LINE_STATUS_REVOKE
    evt.error_code = ""
    evt.fail_sub_type = ""
    evt.is_final_result = False
    evt.is_revoked = True

    events.append(evt)


def _new_common_process_event(
    plan: StaticPlan, row_index: int, backtrack_count: int
) -> AllocationEvent:
    """过程事件公共字段（对应 VBA BT_FillCommonProcessEvent）。"""
    line = plan.lines[row_index]
    return AllocationEvent(
        shipment_no=plan.shipment_no,
        sku=plan.sku,
        wms_order_no=line.wms_order_no,
        line_no=line.line_no,
        demand_d=line.qty,
        process_order=str(row_index + 1),
        dynamic_next_min_qty=_format_next_min_qty(plan, row_index),
        candidate_qc_count="",
        excluded_qc_list="",
        strategy_used="",
        used_qc="",
        qc_before="-",
        qc_after="-",
        lot_expiry_combo_count="",
        is_backtrack_retry="是" if backtrack_count > 0 else "否",
        backtrack_no=backtrack_count,
        line_status="",
        error_code="",
        fail_sub_type="",
        is_final_result=False,
        is_revoked=False,
    )


def _fill_success_debug_fields(
    evt: AllocationEvent,
    result: GroupAllocResult,
    line,
    backtrack_count: int,
) -> None:
    """成功行的最终结果事件字段（对应 VBA BT_FillSuccessDebugFields）。"""
    detail = _find_detail_by_line_no(result.details, line.line_no)
    if detail is not None:
        evt.strategy_used = detail.strategy_used
        evt.used_qc = detail.qc
        evt.qc_before = "-"
        evt.qc_after = "-"
        # VBA 此处硬编码 "1"（即使策略三多组合行），原样保留
        evt.lot_expiry_combo_count = "1"
        evt.line_status = detail.line_status
        evt.error_code = ""
        evt.fail_sub_type = ""
    else:
        evt.strategy_used = "-"
        evt.used_qc = "-"
        evt.qc_before = "-"
        evt.qc_after = "-"
        evt.lot_expiry_combo_count = "-"
        evt.line_status = STATUS_BATCH_IMPORT

    evt.is_backtrack_retry = "是" if backtrack_count > 0 else "否"


def _fill_failure_debug_fields(
    evt: AllocationEvent,
    error_code: str,
    fail_sub_type: str,
    row_index: int,
    first_fail_row: int,
    backtrack_count: int,
    cross_sku_short: bool,
) -> None:
    """失败行的最终结果事件字段（对应 VBA BT_FillFailureDebugFields）。

    row_index / first_fail_row 均为 0-based（VBA 为 1-based，分支条件已换算）。
    """
    evt.line_status = LINE_STATUS_FAILED
    evt.strategy_used = "-"
    evt.used_qc = "-"
    evt.qc_before = "-"
    evt.qc_after = "-"
    evt.lot_expiry_combo_count = "-"

    if cross_sku_short:
        evt.error_code = error_code
        evt.fail_sub_type = fail_sub_type
        evt.process_order = "-"
        evt.candidate_qc_count = "-"
        evt.is_backtrack_retry = "-"
        return

    if fail_sub_type and "预检测" in fail_sub_type:
        evt.error_code = ERR_E09
        evt.fail_sub_type = fail_sub_type
        if int(evt.candidate_qc_count) == 0:
            evt.fail_sub_type = FAIL_SUB_PRECHECK_A
    elif row_index >= first_fail_row and row_index > 0:
        # VBA 1-based 条件 rowIndex >= firstFailRow And rowIndex > 1
        evt.error_code = ERR_E09
        evt.fail_sub_type = FAIL_SUB_CASCADE_SAME_SKU
        evt.process_order = "-"
        evt.candidate_qc_count = "-"
    elif row_index == first_fail_row:
        evt.error_code = error_code or ERR_E09
        evt.fail_sub_type = fail_sub_type or FAIL_SUB_NO_AVAILABLE_QC
    else:
        evt.error_code = error_code
        evt.fail_sub_type = fail_sub_type

    evt.is_backtrack_retry = "是" if backtrack_count > 0 else "否"


def _resolve_fail_sub_type(
    precheck_hit: str, error_code: str, cross_sku_short: bool
) -> str:
    """分配失败子类型归一化（对应 VBA BT_ResolveFailSubType）。"""
    if cross_sku_short:
        return FAIL_SUB_CASCADE_CROSS_SKU
    if precheck_hit == PRECHECK_HIT_A:
        return FAIL_SUB_PRECHECK_A
    if precheck_hit == PRECHECK_HIT_B:
        return FAIL_SUB_PRECHECK_B
    if error_code in (ERR_E10, ERR_E09):
        return FAIL_SUB_PATH_EXHAUSTED
    return ""


def _find_first_fail_row(
    plan: StaticPlan, success: bool, error_code: str, precheck_hit: str
) -> int:
    """首个失败行下标（对应 VBA BT_FindFirstFailRow，返回 0-based）。"""
    row_count = plan.row_count
    if success:
        return row_count  # 无失败行（对应 VBA rowCount + 1）

    if precheck_hit:
        for r, line in enumerate(plan.lines):
            if line.init_qc_count == 0:
                return r
        return 0

    # 动态失败：VBA 硬编码第 2 行（1-based），单行组则为第 1 行
    return 1 if row_count >= 2 else 0


def _find_detail_by_line_no(
    details: list[AllocationDetail], line_no: str
) -> AllocationDetail | None:
    """按行号首个匹配定位明细（对应 VBA BT_FindDetailIndexByLineNo）。"""
    for detail in details:
        if detail.line_no == line_no:
            return detail
    return None


def _format_next_min_qty(plan: StaticPlan, row_index: int) -> str:
    """动态 nextMinQty 显示值（对应 VBA BT_FormatNextMinQty）。

    取排序位置在 row_index 之后各行需求量的最小值；最后一行为空串。
    """
    qtys_after = [line.qty for line in plan.lines[row_index + 1 :]]
    return str(min(qtys_after)) if qtys_after else ""


def _attempt_qc_list(attempt: AllocationAttempt) -> str:
    """分配尝试涉及的 QC 逗号串（对应 VBA BT_AttemptQCList，无明细时为 "-"）。"""
    if not attempt.details:
        return "-"
    return ",".join(d.qc for d in attempt.details)
