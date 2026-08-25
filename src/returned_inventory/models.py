"""M01 基础数据模型（对应 VBA modTypes.bas）。

全系统共用的常量与数据结构定义，不包含任何业务逻辑。
数据生命周期分四层：原始层（excel_input 输出）→ 标准化层（normalize 输出）
→ 问题记录层 → 领域层（库存/分配/校验/运行统计）。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

# -----------------------------------------------------------------------------
# 1. 通用常量
# -----------------------------------------------------------------------------

# 效期单元格类型。excel_input 通过单元格值的 Python 类型判断后写入 RawInventoryRow.expiry_cell_kind。
# 对应 VBA 的 VarType(cell.Value)：datetime ↔ vbDate(7)，str ↔ vbString(8)，None ↔ vbEmpty(0)。
CELL_KIND_EXCEL_DATE = "ExcelDate"
CELL_KIND_TEXT = "TextValue"
CELL_KIND_BLANK = "Blank"
CELL_KIND_OTHER = "Other"

# 来源表名称。用于 FieldNormalizeIssue / ValidationIssue 等问题定位。
SOURCE_RETURN_TABLE = "退单表"
SOURCE_INVENTORY_TABLE = "质检库存表"

# 字段标准化问题类型。
ISSUE_KIND_EMPTY = "Empty"
ISSUE_KIND_FORMAT_ERROR = "FormatError"
ISSUE_KIND_RANGE_ERROR = "RangeError"

# 调试日志级别。生产配置和测试配置都统一使用三档。
DEBUG_LEVEL_OFF = "关闭"
DEBUG_LEVEL_SIMPLE = "简版"
DEBUG_LEVEL_DETAIL = "详细"

# 批号比较模式。
LOT_MODE_INSENSITIVE = "不敏感"
LOT_MODE_SENSITIVE = "敏感"

# QC 情况合法值。
QC_ZP = "ZP"
QC_QC = "QC"
QC_NG = "NG"

# 默认配置值。
DEFAULT_MAX_BACKTRACK_COUNT = 200
DEFAULT_DEBUG_LOG_LEVEL = DEBUG_LEVEL_OFF
DEFAULT_DETAILED_LOG_LIMIT = 100000
DEFAULT_LOT_MODE = LOT_MODE_INSENSITIVE
DEFAULT_NO_EXPIRY_SENTINEL = "2099/01/01"

# 分配前校验错误码（validate 产出，对应需求 §4.1）。
ERR_E01 = "E01"
ERR_E02 = "E02"
ERR_E03 = "E03"
ERR_E04 = "E04"
ERR_E05 = "E05"
ERR_E06 = "E06"
ERR_E07 = "E07"
ERR_E08 = "E08"
ERR_E11 = "E11"

# 分配阶段错误码（backtracking 产出）。
# E09：分配前预检测或回溯耗尽选项，确认无法分配。
# E10：回溯次数超过 max_backtrack_count 配置上限。
# E99：工程守卫发现库存守恒等式被破坏，立即停止运行。
ERR_E09 = "E09"
ERR_E10 = "E10"
ERR_E99 = "E99"

# 汇总表/异常明细占位符。
NA_PLACEHOLDER = "[N/A]"

# 行/退单号状态（status 产出，供 post_validate/output_builder 写入输出表）。
STATUS_BATCH_IMPORT = "批量导入"
STATUS_MANUAL = "手工操作"
STATUS_UNALLOCATED = "无法分配"
LINE_STATUS_FAILED = "分配失败"

# backtracking 短路后未实际分配的 SKU 组标记；status 用于区分"直接失败"与"连带回滚"。
ERROR_CASCADE_ROLLBACK = "连带回滚"

# QC 优先级（数值越小越优先）。仅作平局时的确定性兜底，无业务含义（需求 §6.3.2）。
QC_PRIORITY = {QC_ZP: 1, QC_QC: 2, QC_NG: 3}

# 策略名称（调试日志"使用策略"列）。
STRATEGY_ONE = "策略一"
STRATEGY_TWO = "策略二"
STRATEGY_THREE = "策略三"
STRATEGY_FAILED = "失败"


# -----------------------------------------------------------------------------
# 2. 原始层：excel_input 数据加载输出
# -----------------------------------------------------------------------------


@dataclass
class RawReturnRow:
    """退单表原始行。字段保留 Excel 原始值（如行号可能被 Excel 读成数值 1）。"""

    excel_row_num: int
    shipment_no: Any
    wms_order_no: Any
    sku: Any
    line_no: Any
    qty: Any


@dataclass
class RawInventoryRow:
    """质检库存表原始行。expiry 保留原始单元格值，expiry_cell_kind 记录存储类型。"""

    excel_row_num: int
    shipment_no: Any
    sku: Any
    qc: Any
    lot_no: Any
    expiry: Any
    expiry_cell_kind: str
    qty: Any


# -----------------------------------------------------------------------------
# 3. 标准化层：normalize 数据标准化输出
# -----------------------------------------------------------------------------


@dataclass
class NormalizedReturnLine:
    """标准化后的退单行。valid 字段只说明字段本身是否合法，错误码由 validate 统一生成。"""

    excel_row_num: int
    shipment_no: str
    wms_order_no: str
    sku: str
    line_no: str  # 五位前导零文本，如 "00001"
    qty: int
    line_no_valid: bool
    qty_valid: bool
    empty_fields: str  # 逗号分隔的空字段名，如 "SKU,数量"


@dataclass
class NormalizedInventoryLine:
    """标准化后的质检库存行。"""

    excel_row_num: int
    shipment_no: str
    sku: str
    qc: str  # 已 strip + upper
    lot_no: str  # 已标准化
    expiry: str  # 统一为 YYYY/MM/DD
    qty: int
    qc_valid: bool
    expiry_valid: bool
    qty_valid: bool
    empty_fields: str


@dataclass
class FieldNormalizeIssue:
    """单字段标准化问题记录。normalize 产出，validate 再转换成 E01~E05 等正式错误码。"""

    excel_row_num: int
    source_table: str
    field_name: str
    raw_value: str  # 原始值的字符串表示（供需求 §5.4 数据异常明细表输出）
    issue_kind: str  # ISSUE_KIND_*


# -----------------------------------------------------------------------------
# 4. 领域层：库存、分配、日志、校验和运行统计
# -----------------------------------------------------------------------------


@dataclass(frozen=True)
class InventoryKey:
    """库存五元组键：物流单号 + SKU + QC + 批号 + 效期。"""

    shipment_no: str
    sku: str
    qc: str
    lot_no: str
    expiry: str


@dataclass
class InventoryRow:
    """库存账本中的单行数据。

    original_qty：建账本时的原始数量，整个分配过程中不变，供 guards 守恒断言使用。
    current_qty：当前可用数量，随 deduct/undo 操作变化。
    """

    shipment_no: str
    sku: str
    qc: str
    lot_no: str
    expiry: str
    original_qty: int
    current_qty: int

    @property
    def key(self) -> InventoryKey:
        return InventoryKey(self.shipment_no, self.sku, self.qc, self.lot_no, self.expiry)


# M07 候选库存行与 InventoryRow 结构一致，语义上代表"当前退单行此刻可选用的库存格"。
CandidateRow = InventoryRow


@dataclass
class AllocationDetail:
    """单条分配明细。"""

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
    strategy_used: str = ""


@dataclass
class GroupStats:
    """单个 SKU 组的回溯统计。即使调试日志关闭，也要保留统计值用于运行历史。"""

    shipment_no: str
    sku: str
    backtrack_count: int
    precheck_hit: str  # 空 / "预检测A" / "预检测B"


@dataclass
class AllocationEvent:
    """调试日志事件（19 列输出，见《调试日志19列规格说明.md》）。

    is_final_result=True 的行在「简版」模式下输出；「详细」模式输出全部事件。
    """

    shipment_no: str
    sku: str
    wms_order_no: str
    line_no: str
    demand_d: int
    process_order: str
    dynamic_next_min_qty: str
    candidate_qc_count: str
    excluded_qc_list: str
    strategy_used: str
    used_qc: str
    qc_before: str
    qc_after: str
    lot_expiry_combo_count: str
    is_backtrack_retry: str
    backtrack_no: int
    line_status: str
    error_code: str
    fail_sub_type: str
    is_final_result: bool
    is_revoked: bool


@dataclass
class GroupAllocResult:
    """单个 SKU 组的分配结果。"""

    shipment_no: str
    sku: str
    success: bool
    error_code: str  # E09 / E10 / E99 / 连带回滚 / 空
    stats: GroupStats
    details: list[AllocationDetail] = field(default_factory=list)
    events: list[AllocationEvent] = field(default_factory=list)


@dataclass
class ShipmentAllocResult:
    """单个物流单号的分配结果，含所有 SKU 组结果（含连带回滚）。"""

    shipment_no: str
    group_results: list[GroupAllocResult] = field(default_factory=list)


@dataclass
class ValidationIssue:
    """校验问题。validate 产出，供 status/output_builder 汇总与异常明细使用。"""

    shipment_no: str
    wms_order_no: str
    sku: str
    error_code: str
    source_table: str
    excel_row_num: int
    field_name: str
    raw_value: str
    reason: str


@dataclass
class ValidationResult:
    """校验汇总结果。"""

    has_failures: bool
    failed_shipment_count: int
    issues: list[ValidationIssue] = field(default_factory=list)


@dataclass
class AnomalyRow:
    """数据异常明细行（需求 §5.4）。由 validate.build_anomaly_rows 产出。"""

    source_table: str
    excel_row_num: int
    shipment_no: str
    wms_order_no: str
    sku: str
    field_name: str
    raw_value: str
    error_code: str
    reason: str


@dataclass
class Config:
    """配置结构（对应 VBA ConfigStruct）。"""

    max_backtrack_count: int = DEFAULT_MAX_BACKTRACK_COUNT
    debug_log_level: str = DEFAULT_DEBUG_LOG_LEVEL
    detailed_log_limit: int = DEFAULT_DETAILED_LOG_LIMIT
    lot_case_sensitive: bool = False
    no_expiry_sentinel: str = DEFAULT_NO_EXPIRY_SENTINEL


@dataclass
class RunStats:
    """本次运行汇总统计。runner.build_run_stats 统一构造，Dry Run 时分配相关字段为 0。"""

    total_backtrack_count: int = 0
    max_group_backtrack: int = 0
    validation_fail_count: int = 0
    alloc_success_count: int = 0
    alloc_fail_count: int = 0
    input_return_rows: int = 0
    input_inventory_rows: int = 0
    input_shipment_count: int = 0  # 两表去重合并后的物流单号总数（需求 §5.6）


@dataclass
class WMSStatusEntry:
    """汇总表单条记录（status.aggregate_wms_status 的输出元素）。"""

    shipment_no: str
    wms_order_no: str
    status: str
    reason: str


# M13 输出构建统一行结构：不同输出表列数不同，用 list 承载整行值，按目标表头顺序填充。
OutputRow = list[Any]


@dataclass
class PrecheckResult:
    """sort_filter.run_precheck 的预检测结论。任一字段为 True 代表该 SKU 组分配前已确定失败（E09）。"""

    precheck_a_hit: bool = False  # 预检测A：排序后某行初始可用QC数=0
    precheck_b_hit: bool = False  # 预检测B：多行强制竞争同一QC，合计需求超出库存
