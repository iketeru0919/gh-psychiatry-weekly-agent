Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# ==============================================================
# 延べ日数シート レイアウト確認ツール（読み取り専用）
#
# 各実績記録表の「延べ日数」シートについて、区分別ブロック
# （既定では K5=利用回数合計 / L4:R4=区分 / L5:R5=利用回数）が
# 実際にどこに入っているかを、生の値のまま一覧に出します。
#
# ・ファイルは一切変更しません（開かずにZIPとして読むだけ）
# ・Excelも起動しません
# ・「区分6」がどの列にあるかを自動で探して報告します
# ・数式なのに計算結果が保存されていないセルを検出します
#
# 出力: 出力\延べ日数レイアウト確認.csv
# ==============================================================

$RootPath = "C:\Users\k-takayama\Bihonest Corp. Dropbox\17.ラシエル\00.ラシエル_General\00 実績記録表\実績記録表(請求用)\実績記録表　保存先"

# 空欄なら実行時に入力します。例: "2026年7月"
$TargetMonth = ""

$TargetSheetNames = @("延べ日数", "延日数")
$OutputDir = Join-Path $PSScriptRoot "出力"

# ダンプする範囲
$DumpColumns = @("I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T")
$DumpRows = @(3, 4, 5, 6)

# ==============================================================
# 共通
# ==============================================================
function To-Text {
    param($Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Normalize-Text {
    param($Value)
    $text = To-Text $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    return $text.Normalize([System.Text.NormalizationForm]::FormKC)
}

function Get-ColumnLetter {
    param([int]$ColumnIndex)
    $letter = ""
    $n = $ColumnIndex
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $letter = [char](65 + $rem) + $letter
        $n = [int](($n - $rem - 1) / 26)
    }
    return $letter
}

function Get-MonthKeyFromText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = Normalize-Text $Text
    $y = 0; $m = 0
    foreach ($pattern in @("(\d{4})\s*年\s*(\d{1,2})\s*月", "(\d{4})[\/\.-](\d{1,2})", "^(\d{4})(\d{2})$")) {
        $match = [regex]::Match($t, $pattern)
        if ($match.Success) {
            if ([int]::TryParse($match.Groups[1].Value, [ref]$y) -and [int]::TryParse($match.Groups[2].Value, [ref]$m) -and $m -ge 1 -and $m -le 12) {
                return "{0:D4}-{1:D2}" -f $y, $m
            }
            return $null
        }
    }
    return $null
}

function Format-MonthLabel {
    param([string]$MonthKey)
    if ($MonthKey -match "^(\d{4})-(\d{2})$") { return "{0}年{1}月" -f $Matches[1], [int]$Matches[2] }
    return $MonthKey
}

function Get-MonthFolderInfo {
    param([System.IO.DirectoryInfo]$Directory)

    $monthKey = Get-MonthKeyFromText $Directory.Name
    if ($null -ne $monthKey) {
        return [PSCustomObject]@{ Path = $Directory.FullName; MonthKey = $monthKey }
    }
    $parent = $Directory.Parent
    if ($null -eq $parent) { return $null }
    $parentMatch = [regex]::Match((Normalize-Text $parent.Name), "^(\d{4})年?$")
    $childMatch = [regex]::Match((Normalize-Text $Directory.Name), "^(\d{1,2})月(分)?$")
    $y = 0; $m = 0
    if ($parentMatch.Success -and $childMatch.Success) {
        if ([int]::TryParse($parentMatch.Groups[1].Value, [ref]$y) -and [int]::TryParse($childMatch.Groups[1].Value, [ref]$m) -and $m -ge 1 -and $m -le 12) {
            return [PSCustomObject]@{ Path = $Directory.FullName; MonthKey = "{0:D4}-{1:D2}" -f $y, $m }
        }
    }
    return $null
}

function Get-FileTypeSimple {
    param([string]$FileName)
    $n = Normalize-Text $FileName
    if ($n -match "短期") { return "短期" }
    if ($n -match "3F") { return "入居3F" }
    if ($n -match "2F") { return "入居2F" }
    if ($n -match "1F") { return "入居1F" }
    if ($n -match "日中サービス支援") { return "入居その他" }
    if ($n -match "介護包括") { return "入居その他" }
    return "入居(階不明)"
}

function Is-RecordFile {
    param([string]$FileName)
    $n = Normalize-Text $FileName
    foreach ($ng in @("受給者証", "職員名簿", "家賃", "医療連携", "届出書", "請求金額", "業務日誌", "注意点", "変更点", "ブランク")) {
        if ($n -match $ng) { return $false }
    }
    if ($n -match "様式" -and -not ($n -match "新様式" -and $n -match "実績記録")) { return $false }
    return ($n -match "実績記録" -or $n -match "日中サービス支援" -or $n -match "介護包括" -or $n -match "短期")
}

# ==============================================================
# OpenXML 読み取り（数式のキャッシュ有無も記録）
# ==============================================================
function Get-ZipEntryText {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $stream = $null; $reader = $null
    try {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        return $reader.ReadToEnd()
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Read-SheetForDump {
    param([string]$Path, [string[]]$SheetNames)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = $null; $fileStream = $null
    try {
        $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Read)

        $sharedStrings = @()
        $ssText = Get-ZipEntryText -Zip $zip -EntryName "xl/sharedStrings.xml"
        if (-not [string]::IsNullOrWhiteSpace($ssText)) {
            [xml]$ssXml = $ssText
            foreach ($node in $ssXml.SelectNodes("//*[local-name()='si']")) { $sharedStrings += $node.InnerText }
        }

        $wbText = Get-ZipEntryText -Zip $zip -EntryName "xl/workbook.xml"
        $relText = Get-ZipEntryText -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
        if ([string]::IsNullOrWhiteSpace($wbText) -or [string]::IsNullOrWhiteSpace($relText)) { throw "ブック構造を読み取れません。" }

        [xml]$wbXml = $wbText
        [xml]$relXml = $relText

        # 「開いたときに全再計算」フラグ = 保存済みの値が信用できない印
        $fullCalcOnLoad = $false
        $calcNode = $wbXml.SelectSingleNode("//*[local-name()='calcPr']")
        if ($null -ne $calcNode -and $calcNode.GetAttribute("fullCalcOnLoad") -eq "1") { $fullCalcOnLoad = $true }

        $targets = @{}
        foreach ($node in $relXml.SelectNodes("//*[local-name()='Relationship']")) { $targets[$node.GetAttribute("Id")] = $node.GetAttribute("Target") }

        $sheetNode = $null
        $sheetNameFound = ""
        foreach ($wanted in $SheetNames) {
            foreach ($node in $wbXml.SelectNodes("//*[local-name()='sheet']")) {
                if ((To-Text $node.name) -eq $wanted) { $sheetNode = $node; $sheetNameFound = To-Text $node.name; break }
            }
            if ($null -ne $sheetNode) { break }
        }
        if ($null -eq $sheetNode) {
            foreach ($wanted in $SheetNames) {
                foreach ($node in $wbXml.SelectNodes("//*[local-name()='sheet']")) {
                    if ((To-Text $node.name) -like "*$wanted*") { $sheetNode = $node; $sheetNameFound = To-Text $node.name; break }
                }
                if ($null -ne $sheetNode) { break }
            }
        }
        if ($null -eq $sheetNode) { throw "対象シートが見つかりません。" }

        $rid = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if (-not $targets.ContainsKey($rid)) { throw "対象シートの参照先を読み取れません。" }
        $sheetPath = $targets[$rid].Replace("\", "/").TrimStart("/")
        if ($sheetPath -notlike "xl/*") { $sheetPath = "xl/$sheetPath" }

        $sheetText = Get-ZipEntryText -Zip $zip -EntryName $sheetPath
        if ([string]::IsNullOrWhiteSpace($sheetText)) { throw "対象シートのXMLを読み取れません。" }

        [xml]$sheetXml = $sheetText
        $cells = @{}
        foreach ($cell in $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c']")) {
            $addr = $cell.GetAttribute("r")
            if ([string]::IsNullOrWhiteSpace($addr)) { continue }

            $type = $cell.GetAttribute("t")
            $formulaNode = $cell.SelectSingleNode("*[local-name()='f']")
            $valueNode = $cell.SelectSingleNode("*[local-name()='v']")
            $hasFormula = ($null -ne $formulaNode)

            $display = ""
            if ($type -eq "e") {
                $display = "【数式エラー:" + $(if ($null -ne $valueNode) { $valueNode.InnerText } else { "?" }) + "】"
            } elseif ($type -eq "inlineStr") {
                $inline = $cell.SelectSingleNode("*[local-name()='is']")
                if ($null -ne $inline) { $display = $inline.InnerText }
            } elseif ($null -ne $valueNode) {
                $raw = $valueNode.InnerText
                if ($type -eq "s") {
                    $idx = 0
                    if ([int]::TryParse($raw, [ref]$idx) -and $idx -ge 0 -and $idx -lt $sharedStrings.Count) { $display = $sharedStrings[$idx] }
                } else {
                    $display = $raw
                }
            }

            if ($hasFormula -and [string]::IsNullOrWhiteSpace($display)) {
                $display = "【数式:未計算】"
            }

            $cells[$addr] = $display
        }

        return [PSCustomObject]@{
            SheetName = $sheetNameFound
            Cells = $cells
            FullCalcOnLoad = $fullCalcOnLoad
        }
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
        if ($null -ne $fileStream) { $fileStream.Dispose() }
    }
}

function Get-Cell {
    param([hashtable]$Cells, [string]$Address)
    if ($Cells.ContainsKey($Address)) { return To-Text $Cells[$Address] }
    return ""
}

# 行3〜6のG〜Z列から「6」に見えるセルを探す（区分6がどこにあるかの特定用）
# A〜F列は利用者一覧（利用者名・区分・延べ日数）なので、そちらの「6」を拾わないよう除外します。
function Find-CategorySixCells {
    param([hashtable]$Cells)

    $hits = @()
    for ($row = 3; $row -le 6; $row++) {
        for ($col = 7; $col -le 26; $col++) {
            $letter = Get-ColumnLetter $col
            $text = Normalize-Text (Get-Cell -Cells $Cells -Address ($letter + $row))
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $text = $text -replace "\s+", ""
            $text = $text -replace "^区分", ""
            if ($text -eq "6") { $hits += ($letter + $row) }
        }
    }
    return @($hits)
}

# ==============================================================
# メイン
# ==============================================================
if (-not (Test-Path -LiteralPath $RootPath)) { throw "対象フォルダが見つかりません: $RootPath" }
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

Write-Host "対象ルート: $RootPath"
Write-Host "月フォルダを検索しています..."

$allMonthFolders = @(Get-ChildItem -LiteralPath $RootPath -Directory -Recurse |
    ForEach-Object { Get-MonthFolderInfo $_ } | Where-Object { $null -ne $_ })

if ($allMonthFolders.Count -eq 0) { throw "月フォルダが見つかりません。" }

$monthKeys = @($allMonthFolders | Select-Object -ExpandProperty MonthKey | Sort-Object -Unique)
if ([string]::IsNullOrWhiteSpace($TargetMonth)) {
    Write-Host ""
    Write-Host "見つかった提供月:"
    foreach ($k in $monthKeys) { Write-Host ("  - " + (Format-MonthLabel $k)) }
    Write-Host ""
    $latest = $monthKeys | Sort-Object -Descending | Select-Object -First 1
    $answer = Read-Host ("対象年月を入力してください（例: 2026年7月。空欄なら " + (Format-MonthLabel $latest) + "）")
    if ([string]::IsNullOrWhiteSpace($answer)) { $requestedMonthKey = $latest }
    else {
        $requestedMonthKey = Get-MonthKeyFromText $answer
        if ($null -eq $requestedMonthKey) { throw "年月として読み取れません: $answer" }
    }
} else {
    $requestedMonthKey = Get-MonthKeyFromText $TargetMonth
    if ($null -eq $requestedMonthKey) { throw "TargetMonthを年月として読み取れません: $TargetMonth" }
}

$monthLabel = Format-MonthLabel $requestedMonthKey
Write-Host ("対象月: {0}" -f $monthLabel)

$targetFolders = @($allMonthFolders | Where-Object { $_.MonthKey -eq $requestedMonthKey })
Write-Host ("対象月フォルダ数: {0}" -f $targetFolders.Count)

$results = @()
$fileCount = 0

foreach ($folder in $targetFolders) {
    $relative = $folder.Path.Substring($RootPath.Length).TrimStart("\")
    $facilityFolder = $relative.Split("\")[0]

    $files = @(Get-ChildItem -LiteralPath $folder.Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @(".xlsx", ".xlsm") -and $_.Name -notlike '~$*' -and (Is-RecordFile $_.Name) })

    foreach ($file in $files) {
        $fileCount++
        Write-Host ("  読込中: {0} / {1}" -f $facilityFolder, $file.Name)

        $properties = [ordered]@{
            "フォルダ施設名" = $facilityFolder
            "提供月" = $monthLabel
            "種別" = Get-FileTypeSimple $file.Name
            "ファイル名" = $file.Name
            "更新日時" = $file.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss")
            "競合コピー等" = ""
            "シート名" = ""
            "要再計算" = ""
            "区分6の位置" = ""
            "読取結果" = ""
        }

        $normalizedName = Normalize-Text $file.Name
        $marks = @()
        if ($normalizedName -match "競合コピー|conflicted copy") { $marks += "競合コピー" }
        if ($normalizedName -match "コピー|copy") { $marks += "コピー" }
        if ($normalizedName -match "バックアップ|backup|旧|old") { $marks += "バックアップ/旧" }
        $properties["競合コピー等"] = ($marks -join "/")

        foreach ($row in $DumpRows) {
            foreach ($col in $DumpColumns) { $properties["$col$row"] = "" }
        }

        try {
            $sheet = Read-SheetForDump -Path $file.FullName -SheetNames $TargetSheetNames
            $properties["シート名"] = $sheet.SheetName
            $properties["要再計算"] = $(if ($sheet.FullCalcOnLoad) { "★要再計算（保存値が古い可能性）" } else { "" })

            foreach ($row in $DumpRows) {
                foreach ($col in $DumpColumns) {
                    $properties["$col$row"] = Get-Cell -Cells $sheet.Cells -Address ("$col$row")
                }
            }

            $sixCells = @(Find-CategorySixCells -Cells $sheet.Cells)
            if ($sixCells.Count -eq 0) {
                $properties["区分6の位置"] = "★見つからない"
            } else {
                $properties["区分6の位置"] = ($sixCells -join ", ")
            }

            $properties["読取結果"] = "OK"
        }
        catch {
            $properties["読取結果"] = "エラー: " + $_.Exception.Message
        }

        $results += [PSCustomObject]$properties
    }
}

$csvPath = Join-Path $OutputDir "延べ日数レイアウト確認.csv"
if (Test-Path -LiteralPath $csvPath) { Remove-Item -LiteralPath $csvPath -Force }
$results | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("確認したファイル数: {0}" -f $fileCount)
Write-Host ""

# --- 短期ファイルの要約をコンソールにも出す ---
$shortStay = @($results | Where-Object { $_."種別" -eq "短期" })
if ($shortStay.Count -gt 0) {
    Write-Host "===== 短期ファイルの区分別ブロック（K5=合計 / L4:R4=区分） ====="
    $shortStay |
        Select-Object "フォルダ施設名", "K5", "L4", "M4", "N4", "O4", "P4", "Q4", "R4", "区分6の位置" |
        Format-Table -AutoSize | Out-String -Width 200 | Write-Host

    $noSix = @($shortStay | Where-Object { $_."区分6の位置" -eq "★見つからない" })
    Write-Host ("短期ファイル {0} 件のうち、区分6が見つからないもの: {1} 件" -f $shortStay.Count, $noSix.Count)
    if ($noSix.Count -gt 0) {
        foreach ($r in $noSix) { Write-Host ("  - {0} / {1}" -f $r."フォルダ施設名", $r."ファイル名") }
    }
}

$needRecalc = @($results | Where-Object { -not [string]::IsNullOrWhiteSpace($_."要再計算") })
if ($needRecalc.Count -gt 0) {
    Write-Host ""
    Write-Host ("★ 保存値が古い可能性のあるブック: {0} 件" -f $needRecalc.Count) -ForegroundColor Yellow
    foreach ($r in $needRecalc) { Write-Host ("  - {0} / {1}" -f $r."フォルダ施設名", $r."ファイル名") }
}

$conflicts = @($results | Where-Object { -not [string]::IsNullOrWhiteSpace($_."競合コピー等") })
if ($conflicts.Count -gt 0) {
    Write-Host ""
    Write-Host ("★ 競合コピー／コピー／バックアップの疑いがあるファイル: {0} 件" -f $conflicts.Count) -ForegroundColor Yellow
    foreach ($r in $conflicts) { Write-Host ("  - {0} / {1} [{2}]" -f $r."フォルダ施設名", $r."ファイル名", $r."競合コピー等") }
}

Write-Host ""
Write-Host "出力先:"
Write-Host $csvPath
