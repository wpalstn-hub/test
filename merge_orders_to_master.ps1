param(
    [Parameter(Mandatory = $true)]
    [string]$OrderFolder,

    [Parameter(Mandatory = $true)]
    [string]$MasterWorkbook,

    [string]$MasterSheetName = "",

    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Normalize-Header {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    return (($Text -replace '\s+', '')).Trim().ToLowerInvariant()
}

function Get-HeaderMap {
    param($Worksheet, [int]$HeaderRow, [int]$MaxColumn)

    $map = @{}
    for ($col = 1; $col -le $MaxColumn; $col++) {
        $raw = [string]$Worksheet.Cells.Item($HeaderRow, $col).Text
        $key = Normalize-Header $raw
        if ($key -and -not $map.ContainsKey($key)) {
            $map[$key] = $col
        }
    }
    return $map
}

function Get-LastUsedColumn {
    param($Worksheet)
    return $Worksheet.UsedRange.Columns.Count
}

function Get-LastUsedRow {
    param($Worksheet)
    return $Worksheet.UsedRange.Rows.Count
}

function Find-BestHeaderRow {
    param($Worksheet, $MasterHeaderMap, [int]$MaxScanRows = 10)

    $maxColumn = Get-LastUsedColumn $Worksheet
    $bestRow = 1
    $bestScore = -1

    for ($row = 1; $row -le $MaxScanRows; $row++) {
        $candidateMap = Get-HeaderMap -Worksheet $Worksheet -HeaderRow $row -MaxColumn $maxColumn
        $score = 0
        foreach ($key in $candidateMap.Keys) {
            if ($MasterHeaderMap.ContainsKey($key)) { $score++ }
        }
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestRow = $row
        }
    }

    return $bestRow
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$totalCopied = 0
$processedFiles = 0

try {
    $masterPath = (Resolve-Path -LiteralPath $MasterWorkbook).Path
    $orderPath = (Resolve-Path -LiteralPath $OrderFolder).Path

    $masterWb = $excel.Workbooks.Open($masterPath)
    if ([string]::IsNullOrWhiteSpace($MasterSheetName)) {
        $masterWs = $masterWb.Worksheets.Item(1)
    }
    else {
        $masterWs = $masterWb.Worksheets.Item($MasterSheetName)
    }

    $masterHeaderRow = 1
    $masterMaxCol = Get-LastUsedColumn $masterWs
    $masterHeaderMap = Get-HeaderMap -Worksheet $masterWs -HeaderRow $masterHeaderRow -MaxColumn $masterMaxCol

    if ($masterHeaderMap.Count -eq 0) {
        throw "마스터 시트의 헤더를 읽지 못했습니다. 헤더가 1행에 있는지 확인해 주세요."
    }

    $files = Get-ChildItem -LiteralPath $orderPath -File -Filter '*.xlsx' | Sort-Object Name
    if ($files.Count -eq 0) {
        throw "발주 폴더에서 .xlsx 파일을 찾지 못했습니다: $orderPath"
    }

    foreach ($file in $files) {
        $processedFiles++
        Write-Host "처리 중: $($file.Name)"

        $srcWb = $excel.Workbooks.Open($file.FullName)
        try {
            $srcWs = $srcWb.Worksheets.Item(1)
            $srcMaxCol = Get-LastUsedColumn $srcWs
            $srcLastRow = Get-LastUsedRow $srcWs
            $srcHeaderRow = Find-BestHeaderRow -Worksheet $srcWs -MasterHeaderMap $masterHeaderMap
            $srcHeaderMap = Get-HeaderMap -Worksheet $srcWs -HeaderRow $srcHeaderRow -MaxColumn $srcMaxCol

            $commonHeaders = @($masterHeaderMap.Keys | Where-Object { $srcHeaderMap.ContainsKey($_) })
            if ($commonHeaders.Count -eq 0) {
                Write-Warning "공통 헤더가 없어 건너뜀: $($file.Name)"
                continue
            }

            for ($r = $srcHeaderRow + 1; $r -le $srcLastRow; $r++) {
                $hasData = $false
                foreach ($key in $commonHeaders) {
                    $value = $srcWs.Cells.Item($r, $srcHeaderMap[$key]).Value2
                    if ($null -ne $value -and "${value}".Trim() -ne "") {
                        $hasData = $true
                        break
                    }
                }
                if (-not $hasData) { continue }

                $targetRow = $masterWs.Cells($masterWs.Rows.Count, 1).End(-4162).Row + 1  # xlUp = -4162
                foreach ($key in $commonHeaders) {
                    $srcCol = $srcHeaderMap[$key]
                    $dstCol = $masterHeaderMap[$key]
                    $masterWs.Cells.Item($targetRow, $dstCol).Value2 = $srcWs.Cells.Item($r, $srcCol).Value2
                }
                $totalCopied++
            }
        }
        finally {
            $srcWb.Close($false)
        }
    }

    if (-not $DryRun) {
        $masterWb.Save()
    }

    Write-Host "완료: 파일 $processedFiles개 처리, 행 $totalCopied개 추가"
}
finally {
    if ($masterWb) { $masterWb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
