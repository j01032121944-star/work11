# export_tables.ps1
param(
    [string]$outXlsx = "C:\Users\user\Downloads\work11\2026년_8월_경제활동인구조사_근로형태별_부가조사_통계표.xlsx",
    [string]$csvDir = "C:\Users\user\Downloads\work11\CSV_통계표"
)

$ErrorActionPreference = "Stop"

# Helper to escape XML special characters
function Escape-Xml([string]$str) {
    if ([string]::IsNullOrEmpty($str)) { return "" }
    return $str.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;").Replace("'", "&apos;")
}

# Column letters: 1 -> A, 2 -> B, ..., 26 -> Z, 27 -> AA...
function Get-ColLetter([int]$col) {
    $dividend = $col
    $colName = ""
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $colName = [char](65 + $modulo) + $colName
        $dividend = [int][Math]::Floor(($dividend - $modulo) / 26)
    }
    return $colName
}

# Build Worksheet XML
# colsWidth: array of integers
# rows: array of arrays of cell objects: @{ v = value; t = "str" | "num"; s = styleIndex }
function Build-WorksheetXml([array]$colsWidth, [array]$rows) {
    $sb = New-Object System.Text.StringBuilder
    $null = $sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    $null = $sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    
    if ($colsWidth -and $colsWidth.Count -gt 0) {
        $null = $sb.AppendLine('  <cols>')
        for ($i = 0; $i -lt $colsWidth.Count; $i++) {
            $colNum = $i + 1
            $w = $colsWidth[$i]
            $null = $sb.AppendLine("    <col min=""$colNum"" max=""$colNum"" width=""$w"" customWidth=""1""/>")
        }
        $null = $sb.AppendLine('  </cols>')
    }
    
    $null = $sb.AppendLine('  <sheetData>')
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $rowNum = $r + 1
        $rowCells = $rows[$r]
        if ($null -eq $rowCells -or $rowCells.Count -eq 0) {
            $null = $sb.AppendLine("    <row r=""$rowNum""/>")
            continue
        }
        $null = $sb.AppendLine("    <row r=""$rowNum"">")
        for ($c = 0; $c -lt $rowCells.Count; $c++) {
            $cell = $rowCells[$c]
            if ($null -eq $cell) { continue }
            $colNum = $c + 1
            $cellRef = "$(Get-ColLetter $colNum)$rowNum"
            $s = if ($cell.ContainsKey("s")) { $cell.s } else { 0 }
            $val = $cell.v
            $type = if ($cell.ContainsKey("t")) { $cell.t } else { "str" }
            
            if ($type -eq "num") {
                # Pure numeric cell: Excel treats this as real number!
                $null = $sb.AppendLine("      <c r=""$cellRef"" s=""$s""><v>$val</v></c>")
            } else {
                # Inline string
                $escaped = Escape-Xml "$val"
                $null = $sb.AppendLine("      <c r=""$cellRef"" s=""$s"" t=""inlineStr""><is><t>$escaped</t></is></c>")
            }
        }
        $null = $sb.AppendLine('    </row>')
    }
    $null = $sb.AppendLine('  </sheetData>')
    $null = $sb.AppendLine('</worksheet>')
    return $sb.ToString()
}

# Ensure clean directory
$baseTemp = [System.IO.Path]::GetTempPath()
$randName = "xlsx_build_" + [System.Guid]::NewGuid().ToString("N")
$tempDir = Join-Path -Path $baseTemp -ChildPath $randName
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path -Path $tempDir -ChildPath "_rels") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path -Path $tempDir -ChildPath "xl") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path -Path $tempDir -ChildPath "xl\_rels") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path -Path $tempDir -ChildPath "xl\worksheets") | Out-Null

if (-not (Test-Path $csvDir)) {
    New-Item -ItemType Directory -Force -Path $csvDir | Out-Null
}

# --- STYLES DEFINITION ---
# 0: Normal text (left)
# 1: Big Title (14pt Bold)
# 2: Section Title (11pt Bold)
# 3: Header cell (10pt Bold, Center, Light Gray Fill, Border)
# 4: Integer Number (#,##0, Right, Border) -> CALCULABLE
# 5: Decimal Number (0.0, Right, Border) -> CALCULABLE
# 6: Currency/Decimal with comma (#,##0.0, Right, Border) -> CALCULABLE
# 7: Centered Text (10pt, Center, Border)
# 8: Footnote (9pt, Left, No border)
# 9: Subheader / Category cell (10pt Bold, Left, Border)
$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="0.0"/>
    <numFmt numFmtId="166" formatCode="#,##0.0"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="10"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="11"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="14"/><name val="맑은 고딕"/></font>
    <font><sz val="9"/><color rgb="FF555555"/><name val="맑은 고딕"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE9EEF4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F9FB"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD0D5DD"/></left>
      <right style="thin"><color rgb="FFD0D5DD"/></right>
      <top style="thin"><color rgb="FFD0D5DD"/></top>
      <bottom style="thin"><color rgb="FFD0D5DD"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>
"@
[System.IO.File]::WriteAllText((Join-Path -Path $tempDir -ChildPath "xl\styles.xml"), $stylesXml, [System.Text.Encoding]::UTF8)

# Cell constructor helpers
function C-Str([string]$v, [int]$s = 0) { return @{ v = $v; t = "str"; s = $s } }
function C-Num([double]$v, [int]$s = 4) { return @{ v = $v; t = "num"; s = $s } }
function C-Dec([double]$v, [int]$s = 5) { return @{ v = $v; t = "num"; s = $s } }
function C-DecComma([double]$v, [int]$s = 6) { return @{ v = $v; t = "num"; s = $s } }
function C-Hdr([string]$v) { return @{ v = $v; t = "str"; s = 3 } }
function C-Ctr([string]$v) { return @{ v = $v; t = "str"; s = 7 } }

# -------------------------------------------------------------
# TABLE 1 DATA
# -------------------------------------------------------------
$t1Cols = @(18, 10, 11, 11, 11, 11, 11, 11, 11, 11)

function Get-T1Rows([string]$subTitle, [array]$data, [bool]$includeHeader = $true) {
    $rows = @()
    if ($subTitle) {
        $rows += ,@( (C-Str $subTitle 2) )
    }
    if ($includeHeader) {
        $rows += ,@( (C-Hdr "구분"), (C-Hdr "단위"), (C-Hdr "2019"), (C-Hdr "2020"), (C-Hdr "2021"), (C-Hdr "2022"), (C-Hdr "2023"), (C-Hdr "2024"), (C-Hdr "2025"), (C-Hdr "2026") )
    }
    foreach ($r in $data) {
        $rowCells = @()
        $sIdx = $(if ($r.Bold) { 9 } else { 0 })
        $rowCells += (C-Str $r.Name $sIdx)
        $rowCells += (C-Ctr $r.Unit)
        for ($i = 0; $i -lt 8; $i++) {
            $val = $r.Values[$i]
            if ($r.IsDec) {
                $rowCells += (C-Dec $val 5)
            } else {
                $rowCells += (C-Num $val 4)
            }
        }
        $rows += ,$rowCells
    }
    return $rows
}

$dataTotal = @(
    @{ Name = "임금근로자"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(20559, 20446, 20992, 21724, 21954, 22143, 22413, 22526) },
    @{ Name = "  정규직"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(13078, 13020, 12927, 13568, 13832, 13685, 13845, 13855) },
    @{ Name = "  비정규직"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(7481, 7426, 8066, 8156, 8122, 8459, 8568, 8671) },
    @{ Name = "    비정규직 비중"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(36.4, 36.3, 38.4, 37.5, 37.0, 38.2, 38.2, 38.5) },
    @{ Name = "    한시적"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(4785, 4608, 5171, 5348, 5259, 5628, 5848, 6035) },
    @{ Name = "    시간제"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(3156, 3252, 3512, 3687, 3873, 4256, 4229, 4382) },
    @{ Name = "    비전형"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(2045, 2073, 2278, 2131, 1957, 1903, 1834, 1786) }
)

$dataMale = @(
    @{ Name = "임금근로자"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(11396, 11361, 11517, 11936, 11927, 11908, 11953, 11965) },
    @{ Name = "  정규직"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(8040, 8027, 7941, 8283, 8370, 8293, 8302, 8258) },
    @{ Name = "  비정규직"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(3356, 3335, 3575, 3653, 3557, 3615, 3650, 3708) },
    @{ Name = "    비정규직 비중"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(29.4, 29.4, 31.0, 30.6, 29.8, 30.4, 30.5, 31.0) },
    @{ Name = "    한시적"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(2238, 2079, 2234, 2368, 2349, 2470, 2570, 2706) },
    @{ Name = "    시간제"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(846, 915, 1011, 1125, 1143, 1240, 1187, 1245) },
    @{ Name = "    비전형"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(1125, 1212, 1342, 1209, 1090, 1065, 1006, 942) }
)

$dataFemale = @(
    @{ Name = "임금근로자"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(9163, 9085, 9476, 9788, 10027, 10235, 10460, 10561) },
    @{ Name = "  정규직"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(5038, 4994, 4985, 5285, 5462, 5392, 5543, 5597) },
    @{ Name = "  비정규직"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(4125, 4091, 4491, 4503, 4565, 4844, 4918, 4963) },
    @{ Name = "    비정규직 비중"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(45.0, 45.0, 47.4, 46.0, 45.5, 47.3, 47.0, 47.0) },
    @{ Name = "    한시적"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(2547, 2529, 2938, 2980, 2910, 3158, 3278, 3329) },
    @{ Name = "    시간제"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(2310, 2337, 2502, 2562, 2731, 3016, 3042, 3136) },
    @{ Name = "    비전형"; Unit = "천명"; Bold = $false; IsDec = $false; Values = @(920, 861, 936, 922, 866, 838, 828, 844) }
)

$fn1 = "주 1. ( )안은 임금근로자 중 비정규직이 차지하는 비중임"
$fn2 = "주 2. 근로형태별 비정규직 근로자(한시적, 시간제, 비전형)의 규모는 유형간 중복인원이 포함되기 때문에 합계가 불일치함"

# --- Sheet 1: 통계표1_통합 ---
$rowsT1Combined = @()
$rowsT1Combined += ,@( (C-Str "1. 임금근로자의 근로형태별 규모(8월 기준)" 1) )
$rowsT1Combined += ,@()
$rowsT1Combined += (Get-T1Rows "<전체> (단위: 천명, %)" $dataTotal)
$rowsT1Combined += ,@()
$rowsT1Combined += (Get-T1Rows "<남자> (단위: 천명, %)" $dataMale)
$rowsT1Combined += ,@()
$rowsT1Combined += (Get-T1Rows "<여자> (단위: 천명, %)" $dataFemale)
$rowsT1Combined += ,@()
$rowsT1Combined += ,@( (C-Str $fn1 8) )
$rowsT1Combined += ,@( (C-Str $fn2 8) )

# --- Sheet 2: 통계표1_전체 ---
$rowsT1Total = @()
$rowsT1Total += ,@( (C-Str "1. 임금근로자의 근로형태별 규모(8월 기준) - 전체" 1) )
$rowsT1Total += ,@()
$rowsT1Total += (Get-T1Rows "" $dataTotal)
$rowsT1Total += ,@()
$rowsT1Total += ,@( (C-Str $fn1 8) )
$rowsT1Total += ,@( (C-Str $fn2 8) )

# --- Sheet 3: 통계표1_남자 ---
$rowsT1Male = @()
$rowsT1Male += ,@( (C-Str "1. 임금근로자의 근로형태별 규모(8월 기준) - 남자" 1) )
$rowsT1Male += ,@()
$rowsT1Male += (Get-T1Rows "" $dataMale)
$rowsT1Male += ,@()
$rowsT1Male += ,@( (C-Str $fn1 8) )
$rowsT1Male += ,@( (C-Str $fn2 8) )

# --- Sheet 4: 통계표1_여자 ---
$rowsT1Female = @()
$rowsT1Female += ,@( (C-Str "1. 임금근로자의 근로형태별 규모(8월 기준) - 여자" 1) )
$rowsT1Female += ,@()
$rowsT1Female += (Get-T1Rows "" $dataFemale)
$rowsT1Female += ,@()
$rowsT1Female += ,@( (C-Str $fn1 8) )
$rowsT1Female += ,@( (C-Str $fn2 8) )

# -------------------------------------------------------------
# TABLE 2 DATA (연령계층별, 산업별, 직업별)
# -------------------------------------------------------------
$t2Age = @(
    @{ Name = "비정규직"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(7481, 7426, 8066, 8156, 8122, 8459, 8568, 8671) },
    @{ Name = "15 - 19세"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(2.0, 1.9, 1.8, 1.9, 1.9, 1.7, 1.3, 1.6) },
    @{ Name = "20 - 29세"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(18.2, 17.3, 17.5, 17.3, 17.5, 17.3, 16.7, 15.5) },
    @{ Name = "30 - 39세"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(14.8, 13.8, 12.6, 12.0, 12.2, 12.7, 13.3, 13.6) },
    @{ Name = "40 - 49세"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(18.0, 17.6, 17.6, 16.2, 16.1, 15.5, 14.1, 13.7) },
    @{ Name = "50 - 59세"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(21.0, 20.8, 20.7, 21.1, 20.0, 19.6, 19.1, 19.0) },
    @{ Name = "60세이상"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(25.9, 28.7, 29.8, 31.3, 32.2, 33.2, 35.5, 36.7) }
)

$t2Ind = @(
    @{ Name = "비정규직"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(7481, 7426, 8066, 8156, 8122, 8459, 8568, 8671) },
    @{ Name = "농림어업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(1.0, 1.0, 1.0, 0.8, 0.8, 0.7, 0.7, 0.6) },
    @{ Name = "제조업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(8.9, 8.1, 7.8, 8.1, 8.0, 8.1, 7.7, 7.7) },
    @{ Name = "건설업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(11.4, 11.6, 11.0, 10.4, 9.2, 8.7, 8.0, 7.8) },
    @{ Name = "도소매업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(9.4, 9.7, 9.3, 9.2, 8.7, 8.8, 8.2, 8.1) },
    @{ Name = "운수창고업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(2.9, 3.1, 3.3, 3.1, 3.3, 3.6, 4.0, 4.0) },
    @{ Name = "숙박음식업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(10.2, 9.3, 8.6, 9.5, 9.7, 10.3, 9.5, 9.9) },
    @{ Name = "정보통신업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(1.7, 1.6, 2.3, 2.5, 2.4, 2.4, 2.6, 2.3) },
    @{ Name = "금융보험업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(3.7, 3.4, 3.7, 3.9, 3.8, 3.7, 3.8, 4.1) },
    @{ Name = "부동산업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(2.6, 2.4, 2.6, 2.7, 2.6, 2.2, 2.4, 2.1) },
    @{ Name = "전문과학기술업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(2.6, 2.4, 2.6, 2.3, 2.7, 3.1, 2.9, 2.9) },
    @{ Name = "사업시설지원업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(11.6, 11.1, 11.0, 10.7, 10.7, 10.2, 10.1, 9.8) },
    @{ Name = "공공행정"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(5.0, 5.6, 4.9, 4.7, 4.8, 4.9, 4.6, 4.3) },
    @{ Name = "교육서비스"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(8.4, 7.9, 8.3, 8.6, 8.4, 8.3, 8.3, 8.0) },
    @{ Name = "보건사회복지업"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(13.1, 15.2, 16.8, 17.0, 18.4, 18.3, 20.5, 21.5) },
    @{ Name = "예술스포츠"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(2.6, 2.4, 2.2, 2.3, 2.6, 2.4, 2.4, 2.8) },
    @{ Name = "협회기타서비스"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(3.3, 3.2, 3.0, 3.0, 2.9, 3.2, 3.1, 3.0) },
    @{ Name = "기타"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(1.5, 1.8, 1.5, 1.3, 1.2, 1.1, 1.2, 1.1) }
)

$t2Occ = @(
    @{ Name = "비정규직"; Unit = "천명"; Bold = $true; IsDec = $false; Values = @(7481, 7426, 8066, 8156, 8122, 8459, 8568, 8671) },
    @{ Name = "관리자ㆍ전문가"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(16.2, 15.2, 16.3, 16.6, 17.4, 17.2, 17.7, 17.7) },
    @{ Name = "사무 종사자"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(11.4, 10.6, 10.7, 10.4, 10.6, 10.4, 10.7, 10.4) },
    @{ Name = "서비스·판매 종사자"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(26.0, 25.4, 24.4, 25.4, 26.7, 26.8, 26.3, 27.1) },
    @{ Name = "농림어업 숙련종사자"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(0.5, 0.5, 0.4, 0.4, 0.4, 0.6, 0.5, 0.6) },
    @{ Name = "기능·기계조작 종사자"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(14.9, 14.8, 14.5, 13.8, 13.2, 12.4, 13.0, 12.6) },
    @{ Name = "단순노무 종사자"; Unit = "%"; Bold = $false; IsDec = $true; Values = @(31.0, 33.5, 33.7, 33.3, 31.7, 32.7, 31.9, 31.7) }
)

# --- Sheet 5: 통계표2_연령·산업·직업별 ---
$rowsT2 = @()
$rowsT2 += ,@( (C-Str "2. 비정규직 근로자의 연령계층, 산업, 직업별 규모(8월 기준)" 1) )
$rowsT2 += ,@()
$rowsT2 += (Get-T1Rows "<연령계층별>" $t2Age)
$rowsT2 += ,@()
$rowsT2 += (Get-T1Rows "<산업별>" $t2Ind)
$rowsT2 += ,@()
$rowsT2 += (Get-T1Rows "<직업별>" $t2Occ)
$rowsT2 += ,@()
$rowsT2 += ,@( (C-Str "1. 산업ㆍ직업분류 명칭은 약칭을 사용(부록2 한국표준 산업ㆍ직업분류 참조)" 8) )

# -------------------------------------------------------------
# TABLE 3 DATA (근로조건별 특성)
# -------------------------------------------------------------
$t3Cols = @(20, 14, 10, 11, 11, 11, 11, 11, 11, 11, 11)
$t3Items = @(
    @{ Item = "평균 근속기간"; Unit = "개월"; IsDec = $false;
       Emp = @(71, 72, 70, 72, 74, 76, 78, 80);
       Reg = @(94, 97, 96, 97, 98, 103, 105, 108);
       NonReg = @(29, 29, 29, 30, 32, 34, 35, 36) },
    @{ Item = "평균 취업시간"; Unit = "시간"; IsDec = $true;
       Emp = @(35.9, 37.1, 35.4, 34.7, 33.9, 32.4, 33.4, 36.2);
       Reg = @(38.8, 40.7, 38.7, 37.8, 36.9, 35.4, 36.5, 40.4);
       NonReg = @(30.8, 30.7, 30.2, 29.6, 28.7, 27.6, 28.2, 29.5) },
    @{ Item = "월평균임금"; Unit = "만원"; IsDec = $true; IsCommaDec = $true;
       Emp = @(264.3, 268.1, 273.4, 288.0, 300.7, 312.8, 320.5, 330.5);
       Reg = @(316.5, 323.4, 333.6, 348.0, 362.3, 379.6, 389.6, 405.2);
       NonReg = @(172.9, 171.1, 176.9, 188.1, 195.7, 204.8, 208.8, 211.2) },
    @{ Item = "국민연금 가입률"; Unit = "%"; IsDec = $true;
       Emp = @(69.5, 69.8, 69.4, 70.0, 69.6, 68.8, 68.5, 68.5);
       Reg = @(87.5, 88.0, 88.8, 89.1, 88.0, 88.1, 87.9, 88.6);
       NonReg = @(37.9, 37.8, 38.4, 38.3, 38.4, 37.5, 37.1, 36.5) },
    @{ Item = "건강보험 가입률"; Unit = "%"; IsDec = $true;
       Emp = @(75.7, 76.7, 77.0, 78.5, 78.9, 78.6, 79.0, 79.3);
       Reg = @(91.5, 92.6, 93.6, 94.5, 94.3, 95.0, 95.0, 95.7);
       NonReg = @(48.0, 49.0, 50.3, 51.7, 52.6, 52.2, 53.2, 53.0) },
    @{ Item = "고용보험 가입률"; Unit = "%"; IsDec = $true;
       Emp = @(70.9, 72.6, 75.2, 77.0, 77.0, 77.0, 76.2, 77.0);
       Reg = @(87.2, 89.2, 90.9, 92.2, 91.9, 92.3, 91.8, 93.0);
       NonReg = @(44.9, 46.1, 52.6, 54.0, 54.2, 54.7, 53.7, 54.0) },
    @{ Item = "퇴직급여 수혜율"; Unit = "%"; IsDec = $true;
       Emp = @(74.0, 73.2, 74.0, 74.9, 76.3, 76.8, 76.6, 76.7);
       Reg = @(91.7, 91.9, 93.5, 94.0, 94.5, 95.6, 95.6, 96.2);
       NonReg = @(42.9, 40.4, 42.7, 43.0, 45.3, 46.4, 46.0, 45.7) },
    @{ Item = "상여금 수혜율"; Unit = "%"; IsDec = $true;
       Emp = @(68.8, 68.8, 67.1, 69.9, 70.1, 70.3, 70.2, 70.1);
       Reg = @(86.4, 86.6, 86.7, 88.8, 88.1, 89.0, 89.0, 89.5);
       NonReg = @(38.2, 37.6, 35.7, 38.5, 39.4, 40.1, 39.7, 39.2) },
    @{ Item = "시간외수당 수혜율"; Unit = "%"; IsDec = $true;
       Emp = @(48.2, 50.0, 50.6, 53.0, 54.0, 54.6, 54.6, 55.0);
       Reg = @(60.9, 62.8, 64.9, 67.5, 67.8, 68.7, 69.1, 69.9);
       NonReg = @(25.9, 27.5, 27.6, 28.9, 30.5, 31.8, 31.2, 31.1) },
    @{ Item = "유급휴가 수혜율"; Unit = "%"; IsDec = $true;
       Emp = @(62.9, 64.6, 64.7, 66.3, 67.7, 68.5, 68.8, 69.1);
       Reg = @(80.1, 82.1, 83.3, 84.5, 85.6, 86.9, 87.2, 88.3);
       NonReg = @(33.0, 34.1, 35.1, 35.9, 37.2, 38.7, 39.0, 38.4) }
)

# --- Sheet 6: 통계표3_근로조건별특성 ---
$rowsT3 = @()
$rowsT3 += ,@( (C-Str "3. 임금근로자의 근로조건별 특성(8월 기준)" 1) )
$rowsT3 += ,@()
$rowsT3 += ,@( (C-Hdr "항목"), (C-Hdr "대상구분"), (C-Hdr "단위"), (C-Hdr "2019"), (C-Hdr "2020"), (C-Hdr "2021"), (C-Hdr "2022"), (C-Hdr "2023"), (C-Hdr "2024"), (C-Hdr "2025"), (C-Hdr "2026") )

foreach ($it in $t3Items) {
    # 3 rows per item: 임금근로자, 정규직, 비정규직
    $subList = @(
        @{ Target = "임금근로자"; Bold = $true; Vals = $it.Emp },
        @{ Target = "정규직"; Bold = $false; Vals = $it.Reg },
        @{ Target = "비정규직"; Bold = $false; Vals = $it.NonReg }
    )
    for ($idx = 0; $idx -lt 3; $idx++) {
        $sub = $subList[$idx]
        $rowCells = @()
        $itemStr = $(if ($idx -eq 0) { $it.Item } else { "" })
        $itemStyle = $(if ($idx -eq 0) { 9 } else { 0 })
        $rowCells += (C-Str $itemStr $itemStyle)
        $subStyle = $(if ($sub.Bold) { 9 } else { 0 })
        $rowCells += (C-Str $sub.Target $subStyle)
        $rowCells += (C-Ctr $it.Unit)
        for ($k = 0; $k -lt 8; $k++) {
            $v = $sub.Vals[$k]
            if ($it.IsCommaDec) {
                $rowCells += (C-DecComma $v 6)
            } elseif ($it.IsDec) {
                $rowCells += (C-Dec $v 5)
            } else {
                $rowCells += (C-Num $v 4)
            }
        }
        $rowsT3 += ,$rowCells
    }
}
$rowsT3 += ,@()
$rowsT3 += ,@( (C-Str "1. 사회보험(국민연금, 건강보험, 고용보험) 가입률의 경우, 가입대상 여부에 관계없이 가구내 거주하는 모든 임금근로자를 대상으로 조사하기 때문에 실제 가입대상을 기준으로 한 가입률보다 낮음." 8) )
$rowsT3 += ,@( (C-Str "2. 퇴직급여는 퇴직금 또는 퇴직연금 중 하나라도 해당될 경우의 수혜자 기준임" 8) )
$rowsT3 += ,@( (C-Str "3. 취업시간은 조사대상주간에 공휴일이 주중 또는 주말에 포함되는지에 따라 영향을 받음" 8) )

# -------------------------------------------------------------
# TABLE 4 DATA (시도별 비정규직 근로자 규모)
# -------------------------------------------------------------
$t4Cols = @(16, 8, 12, 12, 12, 12, 12, 12, 12, 12, 12)
$t4Data = @(
    @{ City = "전국"; Emp25 = 22413; Reg25 = 13845; Non25 = 8568; Emp26 = 22526; Reg26 = 13855; Non26 = 8671; ChgEmp = 113; ChgReg = 10; ChgNon = 103 },
    @{ City = "서울"; Emp25 = 4375; Reg25 = 2696; Non25 = 1679; Emp26 = 4326; Reg26 = 2672; Non26 = 1654; ChgEmp = -50; ChgReg = -24; ChgNon = -26 },
    @{ City = "전남광주"; Emp25 = 1264; Reg25 = 671; Non25 = 593; Emp26 = 1260; Reg26 = 685; Non26 = 575; ChgEmp = -4; ChgReg = 14; ChgNon = -18 },
    @{ City = "광주"; Emp25 = 633; Reg25 = 354; Non25 = 278; Emp26 = 633; Reg26 = 360; Non26 = 272; ChgEmp = 0; ChgReg = 6; ChgNon = -6 },
    @{ City = "전남"; Emp25 = 631; Reg25 = 316; Non25 = 315; Emp26 = 627; Reg26 = 325; Non26 = 303; ChgEmp = -4; ChgReg = 9; ChgNon = -12 },
    @{ City = "부산"; Emp25 = 1383; Reg25 = 830; Non25 = 553; Emp26 = 1386; Reg26 = 808; Non26 = 578; ChgEmp = 3; ChgReg = -22; ChgNon = 25 },
    @{ City = "대구"; Emp25 = 913; Reg25 = 561; Non25 = 353; Emp26 = 908; Reg26 = 540; Non26 = 368; ChgEmp = -5; ChgReg = -21; ChgNon = 16 },
    @{ City = "인천"; Emp25 = 1384; Reg25 = 851; Non25 = 533; Emp26 = 1429; Reg26 = 847; Non26 = 582; ChgEmp = 45; ChgReg = -3; ChgNon = 48 },
    @{ City = "대전"; Emp25 = 643; Reg25 = 385; Non25 = 258; Emp26 = 648; Reg26 = 371; Non26 = 278; ChgEmp = 6; ChgReg = -15; ChgNon = 20 },
    @{ City = "울산"; Emp25 = 478; Reg25 = 305; Non25 = 173; Emp26 = 485; Reg26 = 290; Non26 = 195; ChgEmp = 7; ChgReg = -15; ChgNon = 22 },
    @{ City = "세종"; Emp25 = 177; Reg25 = 124; Non25 = 53; Emp26 = 183; Reg26 = 126; Non26 = 57; ChgEmp = 6; ChgReg = 2; ChgNon = 4 },
    @{ City = "경기"; Emp25 = 6320; Reg25 = 4180; Non25 = 2140; Emp26 = 6403; Reg26 = 4230; Non26 = 2173; ChgEmp = 83; ChgReg = 50; ChgNon = 33 },
    @{ City = "강원"; Emp25 = 611; Reg25 = 294; Non25 = 318; Emp26 = 621; Reg26 = 317; Non26 = 304; ChgEmp = 10; ChgReg = 24; ChgNon = -14 },
    @{ City = "충북"; Emp25 = 705; Reg25 = 428; Non25 = 277; Emp26 = 706; Reg26 = 425; Non26 = 281; ChgEmp = 2; ChgReg = -2; ChgNon = 4 },
    @{ City = "충남"; Emp25 = 938; Reg25 = 566; Non25 = 373; Emp26 = 975; Reg26 = 608; Non26 = 367; ChgEmp = 37; ChgReg = 42; ChgNon = -6 },
    @{ City = "전북"; Emp25 = 693; Reg25 = 405; Non25 = 288; Emp26 = 685; Reg26 = 397; Non26 = 289; ChgEmp = -8; ChgReg = -8; ChgNon = 0 },
    @{ City = "경북"; Emp25 = 978; Reg25 = 592; Non25 = 386; Emp26 = 960; Reg26 = 565; Non26 = 396; ChgEmp = -18; ChgReg = -27; ChgNon = 10 },
    @{ City = "경남"; Emp25 = 1278; Reg25 = 798; Non25 = 480; Emp26 = 1273; Reg26 = 819; Non26 = 454; ChgEmp = -5; ChgReg = 21; ChgNon = -26 },
    @{ City = "제주"; Emp25 = 274; Reg25 = 162; Non25 = 112; Emp26 = 278; Reg26 = 156; Non26 = 122; ChgEmp = 5; ChgReg = -6; ChgNon = 10 }
)

# --- Sheet 7: 통계표4_시도별규모 ---
$rowsT4 = @()
$rowsT4 += ,@( (C-Str "4. 시도별 비정규직 근로자 규모(8월 기준)" 1) )
$rowsT4 += ,@()
$rowsT4 += ,@( (C-Hdr "시도"), (C-Hdr "단위"), (C-Hdr "2025 임금근로자"), (C-Hdr "2025 정규직"), (C-Hdr "2025 비정규직"), (C-Hdr "2026 임금근로자"), (C-Hdr "2026 정규직"), (C-Hdr "2026 비정규직"), (C-Hdr "증감 임금근로자"), (C-Hdr "증감 정규직"), (C-Hdr "증감 비정규직") )

foreach ($c in $t4Data) {
    $rowCells = @()
    $isBold = ($c.City -eq "전국")
    $cStyle = $(if ($isBold) { 9 } else { 0 })
    $rowCells += (C-Str $c.City $cStyle)
    $rowCells += (C-Ctr "천명")
    $rowCells += (C-Num $c.Emp25 4)
    $rowCells += (C-Num $c.Reg25 4)
    $rowCells += (C-Num $c.Non25 4)
    $rowCells += (C-Num $c.Emp26 4)
    $rowCells += (C-Num $c.Reg26 4)
    $rowCells += (C-Num $c.Non26 4)
    $rowCells += (C-Num $c.ChgEmp 4)
    $rowCells += (C-Num $c.ChgReg 4)
    $rowCells += (C-Num $c.ChgNon 4)
    $rowsT4 += ,$rowCells
}

# -------------------------------------------------------------
# BUILD SHEETS LIST
# -------------------------------------------------------------
$sheets = @(
    @{ Name = "통계표1_통합"; Cols = $t1Cols; Rows = $rowsT1Combined },
    @{ Name = "통계표1_전체"; Cols = $t1Cols; Rows = $rowsT1Total },
    @{ Name = "통계표1_남자"; Cols = $t1Cols; Rows = $rowsT1Male },
    @{ Name = "통계표1_여자"; Cols = $t1Cols; Rows = $rowsT1Female },
    @{ Name = "통계표2_연령·산업·직업별"; Cols = $t1Cols; Rows = $rowsT2 },
    @{ Name = "통계표3_근로조건별특성"; Cols = $t3Cols; Rows = $rowsT3 },
    @{ Name = "통계표4_시도별규모"; Cols = $t4Cols; Rows = $rowsT4 }
)

# Write Worksheet XMLs
for ($i = 0; $i -lt $sheets.Count; $i++) {
    $sNum = $i + 1
    $xmlContent = Build-WorksheetXml $sheets[$i].Cols $sheets[$i].Rows
    $sheetPath = Join-Path -Path $tempDir -ChildPath "xl\worksheets\sheet$sNum.xml"
    [System.IO.File]::WriteAllText($sheetPath, $xmlContent, [System.Text.Encoding]::UTF8)
}

# Build [Content_Types].xml
$ctSb = New-Object System.Text.StringBuilder
$null = $ctSb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
$null = $ctSb.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
$null = $ctSb.AppendLine('  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
$null = $ctSb.AppendLine('  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
$null = $ctSb.AppendLine('  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
$null = $ctSb.AppendLine('  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
for ($i = 0; $i -lt $sheets.Count; $i++) {
    $sNum = $i + 1
    $null = $ctSb.AppendLine("  <Override PartName=""/xl/worksheets/sheet$sNum.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
}
$null = $ctSb.AppendLine('</Types>')
[System.IO.File]::WriteAllText((Join-Path -Path $tempDir -ChildPath "[Content_Types].xml"), $ctSb.ToString(), [System.Text.Encoding]::UTF8)

# Build _rels/.rels
$relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@
[System.IO.File]::WriteAllText((Join-Path -Path $tempDir -ChildPath "_rels\.rels"), $relsXml, [System.Text.Encoding]::UTF8)

# Build xl/_rels/workbook.xml.rels
$wbRelsSb = New-Object System.Text.StringBuilder
$null = $wbRelsSb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
$null = $wbRelsSb.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
for ($i = 0; $i -lt $sheets.Count; $i++) {
    $sNum = $i + 1
    $null = $wbRelsSb.AppendLine("  <Relationship Id=""rId$sNum"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet$sNum.xml""/>")
}
$stylesRid = "rId$($sheets.Count + 1)"
$null = $wbRelsSb.AppendLine("  <Relationship Id=""$stylesRid"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>")
$null = $wbRelsSb.AppendLine('</Relationships>')
[System.IO.File]::WriteAllText((Join-Path -Path $tempDir -ChildPath "xl\_rels\workbook.xml.rels"), $wbRelsSb.ToString(), [System.Text.Encoding]::UTF8)

# Build xl/workbook.xml
$wbSb = New-Object System.Text.StringBuilder
$null = $wbSb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
$null = $wbSb.AppendLine('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
$null = $wbSb.AppendLine('  <sheets>')
for ($i = 0; $i -lt $sheets.Count; $i++) {
    $sNum = $i + 1
    $sName = Escape-Xml $sheets[$i].Name
    $null = $wbSb.AppendLine("    <sheet name=""$sName"" sheetId=""$sNum"" r:id=""rId$sNum""/>")
}
$null = $wbSb.AppendLine('  </sheets>')
$null = $wbSb.AppendLine('</workbook>')
[System.IO.File]::WriteAllText((Join-Path -Path $tempDir -ChildPath "xl\workbook.xml"), $wbSb.ToString(), [System.Text.Encoding]::UTF8)

# Zip directory to XLSX
[System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
if (Test-Path $outXlsx) { Remove-Item -Force $outXlsx }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $outXlsx)
Remove-Item -Recurse -Force $tempDir
Write-Output "Successfully generated Excel file: $outXlsx"

# -------------------------------------------------------------
# EXPORT CSV FILES (UTF-8 with BOM for Excel Korean support)
# -------------------------------------------------------------
function Export-ToCsv([string]$filePath, [array]$data, [string]$title) {
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $lines = [System.Collections.Generic.List[string]]::new()
    if ($title) {
        $lines.Add("# $title")
    }
    $lines.Add("구분,단위,2019,2020,2021,2022,2023,2024,2025,2026")
    foreach ($r in $data) {
        $vals = ($r.Values | ForEach-Object { "$_" }) -join ","
        $nameClean = $r.Name.Trim()
        $lines.Add("$nameClean,$($r.Unit),$vals")
    }
    [System.IO.File]::WriteAllLines($filePath, $lines, $utf8Bom)
}

Export-ToCsv (Join-Path $csvDir "통계표1_전체.csv") $dataTotal "1. 임금근로자의 근로형태별 규모(8월 기준) - 전체"
Export-ToCsv (Join-Path $csvDir "통계표1_남자.csv") $dataMale "1. 임금근로자의 근로형태별 규모(8월 기준) - 남자"
Export-ToCsv (Join-Path $csvDir "통계표1_여자.csv") $dataFemale "1. 임금근로자의 근로형태별 규모(8월 기준) - 여자"

# Combined CSV
$combLines = [System.Collections.Generic.List[string]]::new()
$combLines.Add("# 1. 임금근로자의 근로형태별 규모(8월 기준) - 통합")
$combLines.Add("집단,구분,단위,2019,2020,2021,2022,2023,2024,2025,2026")
foreach ($group in @( @{ G="전체"; D=$dataTotal }, @{ G="남자"; D=$dataMale }, @{ G="여자"; D=$dataFemale } )) {
    foreach ($r in $group.D) {
        $vals = ($r.Values | ForEach-Object { "$_" }) -join ","
        $combLines.Add("$($group.G),$($r.Name.Trim()),$($r.Unit),$vals")
    }
}
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines((Join-Path $csvDir "통계표1_통합.csv"), $combLines, $utf8Bom)

Write-Output "Successfully exported CSV files to: $csvDir"
