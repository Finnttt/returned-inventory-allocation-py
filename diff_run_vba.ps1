# 阶段11对拍：VBA 侧无头运行。
# 对 data/diff/ 下每个 <case>_input.json：
#   1. 复制 base_vba.xlsm（已同步最新 VBA 模块的测试工作簿副本）为 <case>_vba.xlsm
#   2. 清空输入表/输出表/运行历史数据区，重置配置为默认值
#   3. 按 JSON 写入输入三表（配置按参数名覆盖默认值）
#   4. 调用 RunFullAllocationSilent 并保存
# 用法：
#   powershell -ExecutionPolicy Bypass -File .\diff_run_vba.ps1

param(
    [string]$DiffDir = (Join-Path $PSScriptRoot "data\diff"),
    [string]$BaseWorkbook = (Join-Path $PSScriptRoot "data\diff\base_vba.xlsm")
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path -LiteralPath $BaseWorkbook)) { throw "基础工作簿不存在：$BaseWorkbook" }

$jsonFiles = Get-ChildItem -LiteralPath $DiffDir -Filter "*_input.json" | Sort-Object Name
if ($jsonFiles.Count -eq 0) { throw "未找到 *_input.json，请先运行 diff_build_inputs.py" }

# 生产默认配置（与 生成生产工作簿.ps1 / modConfig 默认值一致）
$defaultConfig = @(
    @("最大回溯次数", "200"),
    @("调试日志级别", "关闭"),
    @("详细日志单表上限", "100000"),
    @("批号比较模式", "不敏感"),
    @("无保质期哨兵值", "2099/01/01")
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$prevSec = $excel.AutomationSecurity
$excel.AutomationSecurity = 1

function Clear-DataRows($ws, [int]$lastCol) {
    $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row  # xlUp
    if ($lastRow -ge 2) {
        $ws.Range($ws.Cells(2, 1), $ws.Cells($lastRow, $lastCol)).ClearContents() | Out-Null
    }
}

function Set-Row($ws, [int]$row, [object[]]$values) {
    for ($c = 0; $c -lt $values.Count; $c++) {
        $v = $values[$c]
        if ($null -eq $v) { $ws.Cells($row, $c + 1).Value2 = "" }
        elseif ($v -is [ValueType]) { $ws.Cells($row, $c + 1).Value2 = [double]$v }
        else { $ws.Cells($row, $c + 1).Value2 = [string]$v }
    }
}

try {
    foreach ($jf in $jsonFiles) {
        $case = $jf.BaseName -replace "_input$", ""
        $payload = Get-Content -LiteralPath $jf.FullName -Raw -Encoding UTF8 | ConvertFrom-Json

        $target = Join-Path $DiffDir ($case + "_vba.xlsm")
        Copy-Item -LiteralPath $BaseWorkbook -Destination $target -Force

        $wb = $excel.Workbooks.Open($target, 0, $false)
        try {
            # 清空输入/输出/历史数据区
            Clear-DataRows $wb.Worksheets.Item("输入_退单表") 5
            Clear-DataRows $wb.Worksheets.Item("输入_质检库存表") 6
            Clear-DataRows $wb.Worksheets.Item("分配状态汇总表") 4
            Clear-DataRows $wb.Worksheets.Item("成功分配明细表") 11
            Clear-DataRows $wb.Worksheets.Item("数据异常明细表") 9
            Clear-DataRows $wb.Worksheets.Item("调试日志") 19
            Clear-DataRows $wb.Worksheets.Item("运行历史记录表") 20

            # 重置配置为默认值，再按 JSON 覆盖
            $cfgWs = $wb.Worksheets.Item("输入_配置")
            Clear-DataRows $cfgWs 3
            for ($i = 0; $i -lt $defaultConfig.Count; $i++) {
                Set-Row $cfgWs ($i + 2) $defaultConfig[$i]
            }
            foreach ($pair in $payload.config) {
                $pname = [string]$pair[0]; $pval = [string]$pair[1]
                $found = $false
                $lastCfg = $cfgWs.Cells($cfgWs.Rows.Count, 1).End(-4162).Row
                for ($r = 2; $r -le $lastCfg; $r++) {
                    if ([string]$cfgWs.Cells($r, 1).Value2 -eq $pname) {
                        $cfgWs.Cells($r, 2).Value2 = $pval
                        $found = $true
                        break
                    }
                }
                if (-not $found) { Set-Row $cfgWs ($lastCfg + 1) @($pname, $pval) }
            }

            # 写入输入数据
            $r = 2
            foreach ($rowVals in $payload.returns) {
                Set-Row $wb.Worksheets.Item("输入_退单表") $r $rowVals
                $r++
            }
            $r = 2
            foreach ($rowVals in $payload.inventory) {
                Set-Row $wb.Worksheets.Item("输入_质检库存表") $r $rowVals
                $r++
            }

            $excel.Run("'" + $wb.Name + "'!RunFullAllocationSilent", $wb) | Out-Null
            $wb.Save()
            Write-Output "DONE $case"
        } finally {
            $wb.Close($false)
        }
    }
} finally {
    try { $excel.AutomationSecurity = $prevSec } catch {}
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
}
