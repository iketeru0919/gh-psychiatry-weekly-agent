Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# ==============================
# 設定
# ==============================
$RootPath = "C:\Users\k-takayama\Bihonest Corp. Dropbox\17.ラシエル\00.ラシエル_General\00 実績記録表\実績記録表(請求用)\実績記録表　保存先"

# 空欄なら最新の月フォルダを自動抽出します。
# 指定する場合の例: "2026年6月", "2026/6", "202606"
$TargetMonth = ""

# 空欄なら実行時に対象施設を対話式で選べます（さらに空欄でEnterなら全施設）。
$TargetFacility = ""

$TargetSheetNames = @("延べ日数", "延日数")
$OutputDir = Join-Path $PSScriptRoot "出力"
$MaxUserRowsToScan = 500

# ------------------------------
# 施設定義テーブル
# ------------------------------
# Pattern     : ファイル名／フォルダ名に含まれる施設のキーワード（NFKC正規化後の文字列に対する正規表現）
# Facility    : 集計時に使う施設名（全シート共通）
# SplitUnits  : $true なら棟（ⅰ/ⅱ/北/南 等）ごとに別施設として集計、$false なら1施設に合算
# UnitStyle   : 棟名の表記。"lower"=ⅰ/ⅱ/ⅲ、"upper"=Ⅰ/Ⅱ/Ⅲ（SplitUnits=$true のときの表示に使用）
# DefaultType : ファイル名から階を判定できなかったときに使う種別（空なら「入居その他」）
#
# ここを直せば、施設の追加・棟の分割方針の変更が1か所で完結します。
$FacilityRules = @(
    [PSCustomObject]@{ Pattern = "萩丘";           Facility = "萩丘";           SplitUnits = $true;  UnitStyle = "lower"; DefaultType = "" }
    [PSCustomObject]@{ Pattern = "西浅田";         Facility = "西浅田";         SplitUnits = $false; UnitStyle = "lower"; DefaultType = "" }
    [PSCustomObject]@{ Pattern = "美田園";         Facility = "美田園";         SplitUnits = $false; UnitStyle = "upper"; DefaultType = "" }
    [PSCustomObject]@{ Pattern = "白浜町";         Facility = "白浜町";         SplitUnits = $false; UnitStyle = "upper"; DefaultType = "" }
    [PSCustomObject]@{ Pattern = "岩沼";           Facility = "岩沼";           SplitUnits = $false; UnitStyle = "lower"; DefaultType = "" }
    [PSCustomObject]@{ Pattern = "おきつアリーナ"; Facility = "おきつアリーナ"; SplitUnits = $false; UnitStyle = "upper"; DefaultType = "入居1F" }
)

# 棟の識別子から入居の階種別を決めるルール（施設ごと）。
# 該当しない場合はファイル名中の 1F/2F/3F を見ます。
$UnitFloorRules = @{
    "白浜町" = @{ "Ⅰ" = "入居1F"; "Ⅱ" = "入居2F"; "Ⅲ" = "入居3F" }
    "美田園" = @{ "北" = "入居1F"; "南" = "入居2F" }
    "岩沼"   = @{ "Ⅰ" = "入居1F"; "Ⅱ" = "入居2F" }
    "西浅田" = @{ "Ⅰ" = "入居1F"; "Ⅱ" = "入居2F" }
}

# 種別の内部コード
$TypeShortStay   = "短期"
$TypeAdmission1F = "入居1F"
$TypeAdmission2F = "入居2F"
$TypeAdmission3F = "入居3F"
$TypeAdmissionOther = "入居その他"
$AdmissionTypes = @($TypeAdmission1F, $TypeAdmission2F, $TypeAdmission3F, $TypeAdmissionOther)
$ShortStayTypes = @($TypeShortStay)

# 区分が空欄だった場合に使う列名（合計を元データと一致させるため、捨てずに1列として出します）
$BlankCategoryLabel = "(区分未設定)"
# マトリクスで固定的に使う列名（区分名と衝突したら区分側に接頭辞を付けます）
$ReservedMatrixColumns = @("施設名", "提供月", "合計")

# ==============================
# 共通関数
# ==============================
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

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
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

    $normalizedText = Normalize-Text $Text
    $year = 0
    $month = 0

    $match = [regex]::Match($normalizedText, "(\d{4})\s*年\s*(\d{1,2})\s*月")
    if ($match.Success) {
        if ([int]::TryParse($match.Groups[1].Value, [ref]$year) -and [int]::TryParse($match.Groups[2].Value, [ref]$month) -and $month -ge 1 -and $month -le 12) {
            return "{0:D4}-{1:D2}" -f $year, $month
        }
        return $null
    }

    $match = [regex]::Match($normalizedText, "(\d{4})[\/\.-](\d{1,2})")
    if ($match.Success) {
        if ([int]::TryParse($match.Groups[1].Value, [ref]$year) -and [int]::TryParse($match.Groups[2].Value, [ref]$month) -and $month -ge 1 -and $month -le 12) {
            return "{0:D4}-{1:D2}" -f $year, $month
        }
        return $null
    }

    $match = [regex]::Match($normalizedText, "^(\d{4})(\d{2})$")
    if ($match.Success) {
        if ([int]::TryParse($match.Groups[1].Value, [ref]$year) -and [int]::TryParse($match.Groups[2].Value, [ref]$month) -and $month -ge 1 -and $month -le 12) {
            return "{0:D4}-{1:D2}" -f $year, $month
        }
        return $null
    }

    $match = [regex]::Match($normalizedText, "^(\d{2})(\d{2})$")
    if ($match.Success) {
        if ([int]::TryParse($match.Groups[1].Value, [ref]$year) -and [int]::TryParse($match.Groups[2].Value, [ref]$month) -and $month -ge 1 -and $month -le 12) {
            return "{0:D4}-{1:D2}" -f (2000 + $year), $month
        }
        return $null
    }

    return $null
}

function Get-MonthFolderPriority {
    param([string]$Path)

    $normalizedPath = Normalize-Text $Path
    if ($normalizedPath -match "\\請求用\\") { return 1 }
    if ($normalizedPath -match "\\作成用\\") { return 2 }
    return 9
}

function Get-FolderContentLastWriteTime {
    param([string]$Path)

    $latestFile = Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @(".xlsx", ".xlsm", ".xls") } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -ne $latestFile) { return $latestFile.LastWriteTime }
    return [DateTime]::MinValue
}

function Get-MonthFolderInfo {
    param([System.IO.DirectoryInfo]$Directory)

    $monthKey = Get-MonthKeyFromText $Directory.Name
    if ($null -ne $monthKey) {
        return [PSCustomObject]@{
            Path = $Directory.FullName
            Name = $Directory.Name
            MonthKey = $monthKey
            LastWriteTime = $Directory.LastWriteTime
            Priority = Get-MonthFolderPriority $Directory.FullName
        }
    }

    $parent = $Directory.Parent
    if ($null -eq $parent) { return $null }

    $parentText = Normalize-Text $parent.Name
    $childText = Normalize-Text $Directory.Name
    $parentMatch = [regex]::Match($parentText, "^(\d{4})年?$")
    $childMatch = [regex]::Match($childText, "^(\d{1,2})月(分)?$")
    $year = 0
    $month = 0

    if ($parentMatch.Success -and $childMatch.Success) {
        if ([int]::TryParse($parentMatch.Groups[1].Value, [ref]$year) -and [int]::TryParse($childMatch.Groups[1].Value, [ref]$month) -and $month -ge 1 -and $month -le 12) {
            return [PSCustomObject]@{
                Path = $Directory.FullName
                Name = $Directory.Name
                MonthKey = "{0:D4}-{1:D2}" -f $year, $month
                LastWriteTime = $Directory.LastWriteTime
                Priority = Get-MonthFolderPriority $Directory.FullName
            }
        }
    }

    return $null
}

function Format-MonthLabel {
    param([string]$MonthKey)
    if ($MonthKey -match "^(\d{4})-(\d{2})$") {
        return "{0}年{1}月" -f $Matches[1], [int]$Matches[2]
    }
    return $MonthKey
}

# ==============================
# 施設名・棟・種別の解決（全シート共通の唯一の実装）
# ==============================

# ファイル名／フォルダ名から、施設判定に使う「素の名前」を作ります。
# NFKC正規化 → 装飾・年月・種別語・様式語をすべて除去。
function Get-CleanFacilityText {
    param([string]$Text)

    $name = Normalize-Text $Text
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }

    # 【】で囲まれた部分があればそれを優先（施設名が入っている運用のため）
    if ($name -match "【([^】]+)】") { $name = $Matches[1] }

    $name = $name -replace "[【】\[\]（）\(\)]", ""
    # 年月表記（2026年6月 / 2026年6月分 / 2026-06 / 2026.6 / R6年6月 / 令和6年6月 / 202606）
    $name = $name -replace "(令和|R)\s*\d{1,2}\s*年\s*\d{1,2}\s*月分?", ""
    $name = $name -replace "\d{4}\s*年\s*\d{1,2}\s*月分?", ""
    $name = $name -replace "\d{4}[\/\.-]\d{1,2}", ""
    $name = $name -replace "\d{1,2}\s*月分", ""
    $name = $name -replace "(?<!\d)\d{6}(?!\d)", ""
    # 書類種別の語
    $name = $name -replace "実績記録票.*$", ""
    $name = $name -replace "実績記録表.*$", ""
    $name = $name -replace "実績記録.*$", ""
    $name = $name -replace "新?様式", ""
    $name = $name -replace "請求用", ""
    $name = $name -replace "作成用", ""
    $name = $name -replace "提出用", ""
    # 種別の語（「短期入所」を「入所」の残骸にしないため長い方から）
    $name = $name -replace "短期入所", ""
    $name = $name -replace "短期", ""
    $name = $name -replace "日中サービス支援型?", ""
    $name = $name -replace "介護包括", ""
    $name = $name -replace "(入居)?[123]F", ""
    # 事業所の接頭辞・記号（NFKCで①〜⑳は数字になるため、数字の接頭辞もここで落とす）
    $name = $name -replace "^(RASIEL|RASIE|ラシエル)", ""
    $name = $name -replace "^[\s0-9\.\-_･・,、★☆〇○◎△▲◇◆■□※]+", ""
    $name = $name -replace "[★☆〇○◎△▲◇◆■□※]", ""
    $name = $name -replace "\s+", ""

    return $name.Trim()
}

# 「素の名前」から棟の識別子（ⅰ/ⅱ/ⅲ/北/南）を取り出します。
# NFKC正規化済みなので ⅰ→i、Ⅲ→III になっており、-match は大文字小文字を区別しません。
function Get-UnitToken {
    param(
        [string]$CleanName,
        [string]$FacilityPattern
    )

    $match = [regex]::Match($CleanName, $FacilityPattern + "\s*(iii|ii|i|北|南)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return "" }

    $token = $match.Groups[1].Value
    switch -Regex ($token) {
        "^iii$" { return "Ⅲ" }
        "^ii$"  { return "Ⅱ" }
        "^i$"   { return "Ⅰ" }
        "北"    { return "北" }
        "南"    { return "南" }
    }
    return ""
}

# 棟識別子の表示用表記（内部では Ⅰ/Ⅱ/Ⅲ に統一し、表示だけ施設ごとの流儀に合わせます）。
function Format-UnitToken {
    param(
        [string]$UnitToken,
        [string]$UnitStyle
    )

    if ($UnitStyle -ne "lower") { return $UnitToken }

    switch ($UnitToken) {
        "Ⅰ" { return "ⅰ" }
        "Ⅱ" { return "ⅱ" }
        "Ⅲ" { return "ⅲ" }
    }
    return $UnitToken
}

# ファイル名（＋フォルダ名）から施設・棟を一意に決めます。
# 戻り値: Facility=集計に使う施設名 / Unit=棟キー（同一種別の重複判定に使う） / Matched=定義表に一致したか
function Get-FacilityIdentity {
    param(
        [string]$FileName,
        [string]$FolderFacilityName = ""
    )

    $cleanFromFile = Get-CleanFacilityText $FileName
    $cleanFromFolder = Get-CleanFacilityText $FolderFacilityName

    foreach ($source in @($cleanFromFile, $cleanFromFolder)) {
        if ([string]::IsNullOrWhiteSpace($source)) { continue }

        foreach ($rule in $FacilityRules) {
            if ($source -notmatch $rule.Pattern) { continue }

            # 棟識別子はファイル名側を優先し、無ければフォルダ名側も見る
            $unitToken = Get-UnitToken -CleanName $cleanFromFile -FacilityPattern $rule.Pattern
            if ([string]::IsNullOrWhiteSpace($unitToken) -and -not [string]::IsNullOrWhiteSpace($cleanFromFolder)) {
                $unitToken = Get-UnitToken -CleanName $cleanFromFolder -FacilityPattern $rule.Pattern
            }

            $facility = $rule.Facility
            if ($rule.SplitUnits) {
                if ([string]::IsNullOrWhiteSpace($unitToken)) {
                    # 棟を分けて集計する施設なのに棟が特定できない（例: 萩丘共通の短期ファイル）。
                    # 勝手にどれかの棟に寄せると数値が狂うため、専用の行として分離します。
                    $facility = "{0}(棟共通)" -f $rule.Facility
                } else {
                    $facility = "{0}{1}" -f $rule.Facility, (Format-UnitToken -UnitToken $unitToken -UnitStyle $rule.UnitStyle)
                }
            }

            $unitKey = $rule.Facility
            if (-not [string]::IsNullOrWhiteSpace($unitToken)) {
                $unitKey = "{0}{1}" -f $rule.Facility, $unitToken
            }

            return [PSCustomObject]@{
                Facility     = $facility
                Unit         = $unitKey
                UnitToken    = $unitToken
                RuleFacility = $rule.Facility
                DefaultType  = $rule.DefaultType
                Matched      = $true
            }
        }
    }

    # 定義表に無い施設。ファイル名由来 → フォルダ名由来の順で素の名前をそのまま使います。
    $fallback = $cleanFromFile
    if ([string]::IsNullOrWhiteSpace($fallback)) { $fallback = $cleanFromFolder }
    if ([string]::IsNullOrWhiteSpace($fallback)) { $fallback = To-Text $FileName }

    return [PSCustomObject]@{
        Facility     = $fallback
        Unit         = $fallback
        UnitToken    = ""
        RuleFacility = ""
        DefaultType  = ""
        Matched      = $false
    }
}

# 種別（短期／入居1F/2F/3F/その他）を決めます。
# 「短期」を最優先で判定します（旧版は 1F/2F を先に見ていたため「短期1F」が入居扱いになっていました）。
function Get-FileType {
    param(
        [string]$FileName,
        $Identity = $null
    )

    $normalizedFileName = Normalize-Text $FileName

    if ($normalizedFileName -match "短期") { return $TypeShortStay }
    if ($normalizedFileName -match "3F") { return $TypeAdmission3F }
    if ($normalizedFileName -match "2F") { return $TypeAdmission2F }
    if ($normalizedFileName -match "1F") { return $TypeAdmission1F }
    if ($normalizedFileName -match "日中サービス支援") { return $TypeAdmissionOther }
    if ($normalizedFileName -match "介護包括") { return $TypeAdmissionOther }

    # 棟の識別子から階を決める（白浜町Ⅰ→1F、美田園北→1F など）
    if ($null -ne $Identity -and $Identity.Matched -and -not [string]::IsNullOrWhiteSpace($Identity.UnitToken)) {
        if ($UnitFloorRules.ContainsKey($Identity.RuleFacility)) {
            $floorMap = $UnitFloorRules[$Identity.RuleFacility]
            if ($floorMap.ContainsKey($Identity.UnitToken)) {
                return $floorMap[$Identity.UnitToken]
            }
        }
    }

    # 施設定義テーブルの既定種別（例: おきつアリーナ → 入居1F）
    if ($null -ne $Identity -and $Identity.Matched -and -not [string]::IsNullOrWhiteSpace($Identity.DefaultType)) {
        return $Identity.DefaultType
    }

    # 判定できなくても集計対象からは外しません（旧版は「不明」として丸ごと捨てていました）。
    return $TypeAdmissionOther
}

function Test-IsFloorUndetermined {
    param(
        [string]$FileName,
        $Identity
    )

    $normalizedFileName = Normalize-Text $FileName
    if ($normalizedFileName -match "短期|[123]F|日中サービス支援|介護包括") { return $false }
    if ($null -ne $Identity -and $Identity.Matched) {
        if (-not [string]::IsNullOrWhiteSpace($Identity.UnitToken) -and $UnitFloorRules.ContainsKey($Identity.RuleFacility)) {
            if ($UnitFloorRules[$Identity.RuleFacility].ContainsKey($Identity.UnitToken)) { return $false }
        }
        if (-not [string]::IsNullOrWhiteSpace($Identity.DefaultType)) { return $false }
    }
    return $true
}

function Is-ExcludedExcelFile {
    param([string]$FileName)

    $normalizedFileName = Normalize-Text $FileName

    if ($normalizedFileName -match "受給者証") { return $true }
    if ($normalizedFileName -match "職員名簿") { return $true }
    if ($normalizedFileName -match "家賃") { return $true }
    if ($normalizedFileName -match "医療連携") { return $true }
    if ($normalizedFileName -match "届出書") { return $true }
    if ($normalizedFileName -match "請求金額") { return $true }
    if ($normalizedFileName -match "業務日誌") { return $true }
    if ($normalizedFileName -match "注意点") { return $true }
    if ($normalizedFileName -match "変更点") { return $true }
    if ($normalizedFileName -match "ブランク") { return $true }
    if ($normalizedFileName -match "様式") {
        if ($normalizedFileName -match "新様式" -and $normalizedFileName -match "実績記録") {
            return $false
        }
        return $true
    }

    return $false
}

function Is-RecordCandidateExcelFile {
    param([string]$FileName)

    $normalizedFileName = Normalize-Text $FileName

    if (Is-ExcludedExcelFile $FileName) { return $false }
    if ($normalizedFileName -match "実績記録") { return $true }
    if ($normalizedFileName -match "実績記録票") { return $true }
    if ($normalizedFileName -match "日中サービス支援") { return $true }
    if ($normalizedFileName -match "介護包括") { return $true }
    if ($normalizedFileName -match "短期") { return $true }

    return $false
}

function Get-FacilityName {
    param(
        [string]$Root,
        [string]$MonthFolderPath
    )

    $relative = $MonthFolderPath.Substring($Root.Length).TrimStart("\")
    $firstPart = $relative.Split("\")[0]
    if ([string]::IsNullOrWhiteSpace($firstPart)) { return "" }
    return $firstPart
}

# ==============================
# 数値・区分の正規化
# ==============================
function Convert-ToNumberOrZero {
    param($Value)

    $text = To-Text $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 0 }

    $text = (Normalize-Text $text).Replace(",", "")
    $number = 0.0
    if ([double]::TryParse($text, [ref]$number)) { return $number }

    return 0
}

function Format-NumberForCsv {
    param([double]$Value)

    if ([Math]::Abs($Value - [Math]::Round($Value)) -lt 0.0000001) {
        return ([int][Math]::Round($Value)).ToString()
    }

    return $Value.ToString()
}

# 区分の表記ゆれを吸収します（全角6／区分6／６ → すべて "6"、無し／－ → "なし"）。
function Normalize-CategoryText {
    param($Value)

    $text = Normalize-Text $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $text = $text -replace "\s+", ""
    $text = $text -replace "^区分", ""
    $text = $text -replace "^:", ""

    if ($text -match "^(なし|無し|無|該当なし|非該当|[-‐‑–—―ー])$") { return "なし" }
    if ($text -match "^(\d+)$") { return [string][int]$Matches[1] }

    return $text
}

function Get-CategorySortKey {
    param($Value)

    $text = To-Text $Value
    if ($text -eq "なし") { return 999997 }
    if ($text -eq $BlankCategoryLabel) { return 999999 }

    $number = 0
    if ([int]::TryParse($text, [ref]$number)) {
        return -1 * $number
    }

    return 999998
}

function Get-CategoryTypeSortKey {
    param([string]$Type)

    switch ($Type) {
        "入居1F" { return 1 }
        "入居2F" { return 2 }
        "入居3F" { return 3 }
        "入居その他" { return 4 }
        "入居合計" { return 5 }
        "短期" { return 6 }
        default { return 9 }
    }
}

# 区分名が固定列名（施設名・提供月・合計）と衝突すると列が上書きされて壊れるため、接頭辞で退避します。
function Get-SafeCategoryColumnName {
    param([string]$Category)

    if ($ReservedMatrixColumns -contains $Category) {
        return "区分:{0}" -f $Category
    }
    return $Category
}

# ==============================
# Excel（OpenXML）読み取り
# ==============================
function Get-ZipEntryText {
    param(
        $Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }

    $stream = $null
    $reader = $null
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

function Resolve-OpenXmlTargetPath {
    param([string]$Target)

    $targetPath = $Target.Replace("\", "/").TrimStart("/")
    if ($targetPath -like "xl/*") { return $targetPath }
    return "xl/$targetPath"
}

function Get-OpenXmlSharedStrings {
    param($Zip)

    $sharedStrings = @()
    $sharedStringsXmlText = Get-ZipEntryText -Zip $Zip -EntryName "xl/sharedStrings.xml"
    if (-not [string]::IsNullOrWhiteSpace($sharedStringsXmlText)) {
        [xml]$sharedStringsXml = $sharedStringsXmlText
        $stringNodes = $sharedStringsXml.SelectNodes("//*[local-name()='si']")
        foreach ($node in $stringNodes) {
            $sharedStrings += $node.InnerText
        }
    }
    return $sharedStrings
}

function Convert-OpenXmlSheetXml {
    param(
        [string]$SheetXmlText,
        [object[]]$SharedStrings
    )

    [xml]$sheetXml = $SheetXmlText
    $cells = @{}
    $cellErrors = @{}
    foreach ($cell in $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c']")) {
        $address = $cell.GetAttribute("r")
        if ([string]::IsNullOrWhiteSpace($address)) { continue }

        $value = ""
        $type = $cell.GetAttribute("t")

        if ($type -eq "e") {
            $valueNode = $cell.SelectSingleNode("*[local-name()='v']")
            $errorText = if ($null -ne $valueNode) { $valueNode.InnerText } else { "#ERROR" }
            $cellErrors[$address] = $errorText
            $cells[$address] = ""
            continue
        }

        if ($type -eq "inlineStr") {
            $inlineNode = $cell.SelectSingleNode("*[local-name()='is']")
            if ($null -ne $inlineNode) { $value = $inlineNode.InnerText }
        } else {
            $valueNode = $cell.SelectSingleNode("*[local-name()='v']")
            if ($null -ne $valueNode) {
                $rawValue = $valueNode.InnerText
                if ($type -eq "s") {
                    $index = 0
                    if ([int]::TryParse($rawValue, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) {
                        $value = $SharedStrings[$index]
                    }
                } else {
                    $value = $rawValue
                }
            }
        }

        $cells[$address] = $value
    }

    return [PSCustomObject]@{
        Cells = $cells
        CellErrors = $cellErrors
    }
}

# ブックを1回だけ開いて全シートを読みます（旧版は同じブックを2回開いて2回パースしていました）。
function Read-OpenXmlWorkbookSheets {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = $null
    $fileStream = $null
    try {
        $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Read)

        $sharedStrings = @(Get-OpenXmlSharedStrings -Zip $zip)

        $workbookXmlText = Get-ZipEntryText -Zip $zip -EntryName "xl/workbook.xml"
        $relsXmlText = Get-ZipEntryText -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
        if ([string]::IsNullOrWhiteSpace($workbookXmlText) -or [string]::IsNullOrWhiteSpace($relsXmlText)) {
            throw "Excelブック構造を読み取れません。"
        }

        [xml]$workbookXml = $workbookXmlText
        [xml]$relsXml = $relsXmlText

        $relationshipTargets = @{}
        foreach ($node in $relsXml.SelectNodes("//*[local-name()='Relationship']")) {
            $relationshipTargets[$node.GetAttribute("Id")] = $node.GetAttribute("Target")
        }

        $results = @()
        foreach ($sheetNode in $workbookXml.SelectNodes("//*[local-name()='sheet']")) {
            $sheetName = To-Text $sheetNode.name
            $relationshipId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            if (-not $relationshipTargets.ContainsKey($relationshipId)) { continue }

            $sheetPath = Resolve-OpenXmlTargetPath $relationshipTargets[$relationshipId]
            $sheetXmlText = Get-ZipEntryText -Zip $zip -EntryName $sheetPath
            if ([string]::IsNullOrWhiteSpace($sheetXmlText)) { continue }

            $parsed = Convert-OpenXmlSheetXml -SheetXmlText $sheetXmlText -SharedStrings $sharedStrings
            $results += [PSCustomObject]@{
                SheetName = $sheetName
                Cells = $parsed.Cells
                CellErrors = $parsed.CellErrors
            }
        }

        return @($results)
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
        if ($null -ne $fileStream) { $fileStream.Dispose() }
    }
}

function Select-TargetSheet {
    param(
        [object[]]$Sheets,
        [string[]]$SheetNames
    )

    foreach ($sheetName in $SheetNames) {
        foreach ($sheet in $Sheets) {
            if ((To-Text $sheet.SheetName) -eq $sheetName) { return $sheet }
        }
    }
    foreach ($sheetName in $SheetNames) {
        foreach ($sheet in $Sheets) {
            if ((To-Text $sheet.SheetName) -like "*$sheetName*") { return $sheet }
        }
    }
    return $null
}

function Get-WorksheetByName {
    param(
        $Workbook,
        [string[]]$SheetNames
    )

    foreach ($sheetName in $SheetNames) {
        foreach ($sheet in $Workbook.Worksheets) {
            if ((To-Text $sheet.Name) -eq $sheetName) { return $sheet }
        }
    }
    foreach ($sheetName in $SheetNames) {
        foreach ($sheet in $Workbook.Worksheets) {
            if ((To-Text $sheet.Name) -like "*$sheetName*") { return $sheet }
        }
    }
    return $null
}

function Test-IsExcelErrorText {
    param([string]$Text)
    return $Text -match "^#(REF!|VALUE!|DIV/0!|NAME\?|NULL!|NUM!|N/A)$"
}

function Test-CanReadExcelFile {
    param([string]$Path)

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-CellValueFromMap {
    param(
        [hashtable]$Cells,
        [string]$Address
    )

    if ($Cells.ContainsKey($Address)) { return To-Text $Cells[$Address] }
    return ""
}

function Get-LastDataRowFromMap {
    param(
        [hashtable]$Cells,
        [string[]]$ColumnLetters
    )

    $pattern = "^(" + ($ColumnLetters -join "|") + ")(\d+)$"
    $lastRow = 0
    foreach ($address in $Cells.Keys) {
        if ($address -match $pattern -and -not [string]::IsNullOrWhiteSpace((To-Text $Cells[$address]))) {
            $row = [int]$Matches[2]
            if ($row -gt $lastRow) { $lastRow = $row }
        }
    }
    return $lastRow
}

function Get-LastDataRowInColumns {
    param(
        $Worksheet,
        [string[]]$ColumnLetters
    )

    $lastRow = 0
    foreach ($columnLetter in $ColumnLetters) {
        $range = $null
        $found = $null
        try {
            $range = $Worksheet.Range("${columnLetter}:${columnLetter}")
            $found = $range.Find("*", $range.Cells.Item(1, 1), -4163, 2, 1, 2, $false)
            if ($null -ne $found -and [int]$found.Row -gt $lastRow) {
                $lastRow = [int]$found.Row
            }
        }
        finally {
            Release-ComObject $found
            Release-ComObject $range
        }
    }
    return $lastRow
}

function Find-UserTableColumnsFromMap {
    param(
        [hashtable]$Cells,
        [int]$MaxScanRow = 5,
        [int]$MaxScanCol = 8
    )

    for ($row = 1; $row -le $MaxScanRow; $row++) {
        $nameCol = 0
        $categoryCol = 0
        $daysCol = 0

        for ($col = 1; $col -le $MaxScanCol; $col++) {
            $text = Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $col) + $row)
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            if ($nameCol -eq 0 -and $text -match "利用者|氏名|名前") { $nameCol = $col }
            elseif ($categoryCol -eq 0 -and $text -match "区分") { $categoryCol = $col }
            elseif ($daysCol -eq 0 -and $text -match "延べ日数|延日数") { $daysCol = $col }
        }

        if ($nameCol -gt 0 -and $categoryCol -gt 0 -and $daysCol -gt 0) {
            return [PSCustomObject]@{
                HeaderRow = $row
                NameCol = $nameCol
                CategoryCol = $categoryCol
                DaysCol = $daysCol
            }
        }
    }

    return $null
}

function Find-UserTableColumnsFromWorksheet {
    param(
        $Worksheet,
        [int]$MaxScanRow = 5,
        [int]$MaxScanCol = 8
    )

    for ($row = 1; $row -le $MaxScanRow; $row++) {
        $nameCol = 0
        $categoryCol = 0
        $daysCol = 0

        for ($col = 1; $col -le $MaxScanCol; $col++) {
            $text = To-Text $Worksheet.Cells.Item($row, $col).Value2
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            if ($nameCol -eq 0 -and $text -match "利用者|氏名|名前") { $nameCol = $col }
            elseif ($categoryCol -eq 0 -and $text -match "区分") { $categoryCol = $col }
            elseif ($daysCol -eq 0 -and $text -match "延べ日数|延日数") { $daysCol = $col }
        }

        if ($nameCol -gt 0 -and $categoryCol -gt 0 -and $daysCol -gt 0) {
            return [PSCustomObject]@{
                HeaderRow = $row
                NameCol = $nameCol
                CategoryCol = $categoryCol
                DaysCol = $daysCol
            }
        }
    }

    return $null
}

# 「合計」「小計」等の集計行を利用者として拾わないための判定。
function Test-IsAggregateRowLabel {
    param([string]$Text)

    $normalized = Normalize-Text $Text
    $normalized = $normalized -replace "\s+", ""
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $false }
    return ($normalized -match "^(合計|総合計|小計|計|累計|総計|平均|人数|延べ人数)$")
}

# ==============================
# 利用者個票（実績記録票）の読み取り
# ==============================
function Normalize-LabelText {
    param([string]$Text)

    $normalized = Normalize-Text $Text
    if ([string]::IsNullOrWhiteSpace($normalized)) { return "" }
    return ($normalized -replace "\s+", "")
}

# 個票2行目に現れる見出しの一覧。
# 「見出しの右隣が空欄のとき、さらに右の“別の見出し”を値として拾ってしまう」事故を防ぐために使います。
$UserHeaderLabelPatterns = @(
    "受給者証番号",
    "支給決定障害者|支給決定|^氏名$|フリガナ",
    "^(障害支援)?区分$",
    "部屋番号|居室番号",
    "事業所番号|事業所名|事業者",
    "契約支給量",
    "市町村|市区町村",
    "利用者負担|上限月額",
    "提供年月|年月"
)

function Get-RowLabelItems {
    param(
        [hashtable]$Cells,
        [int]$Row,
        [int]$MaxCol = 30
    )

    $items = @()
    for ($col = 1; $col -le $MaxCol; $col++) {
        $raw = Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $col) + $Row)
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $items += [PSCustomObject]@{
            Col  = $col
            Raw  = To-Text $raw
            Norm = Normalize-LabelText $raw
        }
    }
    return @($items)
}

function Test-IsKnownHeaderLabel {
    param([string]$NormalizedText)

    foreach ($pattern in $UserHeaderLabelPatterns) {
        if ($NormalizedText -match $pattern) { return $true }
    }
    return $false
}

# 見出しに一致したセルの「次の非空セル」を値として返します。
# ただし次の非空セルがそれ自体別の見出しなら、値は未入力とみなして空を返します。
function Find-RowLabelValue {
    param(
        [hashtable]$Cells,
        [int]$Row,
        [string]$LabelPattern,
        [int]$MaxCol = 30,
        [switch]$AllowLabelAsValue
    )

    # 空行だと関数の出力が $null になるため、必ず配列に包み直します（StrictMode対策）
    $items = @(Get-RowLabelItems -Cells $Cells -Row $Row -MaxCol $MaxCol)

    for ($i = 0; $i -lt $items.Count; $i++) {
        if ($items[$i].Norm -notmatch $LabelPattern) { continue }
        if ($i + 1 -ge $items.Count) { return "" }

        $next = $items[$i + 1]
        if (-not $AllowLabelAsValue -and (Test-IsKnownHeaderLabel $next.Norm)) { return "" }
        return $next.Raw
    }

    return ""
}

function Get-UserRecordDetailRow {
    param(
        [string]$FacilityName,
        [string]$MonthLabel,
        [string]$FileType,
        [string]$FileName,
        [string]$SheetName,
        [hashtable]$Cells
    )

    $isRecordSheet = $false
    # 旧版は1行目しか見ていなかったため、タイトルが2〜3行目にある個票を取りこぼしていました。
    for ($row = 1; $row -le 3 -and -not $isRecordSheet; $row++) {
        for ($col = 1; $col -le 30; $col++) {
            $text = Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $col) + $row)
            if ($text -match "実績記録票") { $isRecordSheet = $true; break }
        }
    }
    if (-not $isRecordSheet) { return $null }

    $userName = Find-RowLabelValue -Cells $Cells -Row 2 -LabelPattern "支給決定|^氏名$|氏名"
    if ([string]::IsNullOrWhiteSpace($userName)) { return $null }

    $recipientNumber = Find-RowLabelValue -Cells $Cells -Row 2 -LabelPattern "受給者証番号"
    if ($recipientNumber -match "^\d{1,9}$") {
        $recipientNumber = $recipientNumber.PadLeft(10, "0")
    }
    # 「障害支援区分」のようなラベル表記ゆれに対応し、値側の「区分6」は 6 に正規化します。
    $categoryLevel = Normalize-CategoryText (Find-RowLabelValue -Cells $Cells -Row 2 -LabelPattern "^(障害支援)?区分$" -AllowLabelAsValue)
    $roomNumber = Find-RowLabelValue -Cells $Cells -Row 2 -LabelPattern "部屋番号|居室番号" -AllowLabelAsValue

    $treatmentImprovement = ""
    for ($col = 1; $col -le 30; $col++) {
        $letter = Get-ColumnLetter $col
        $labelText = Normalize-LabelText (Get-CellValueFromMap -Cells $Cells -Address ($letter + "3"))
        if ($labelText -match "処遇改善") {
            $treatmentImprovement = To-Text (Get-CellValueFromMap -Cells $Cells -Address ($letter + "4"))
            break
        }
    }

    $severeSupport1 = Find-RowLabelValue -Cells $Cells -Row 3 -LabelPattern "重度障害者支援加算\(I\)" -AllowLabelAsValue
    $severeSupport2 = Find-RowLabelValue -Cells $Cells -Row 3 -LabelPattern "重度障害者支援加算\(II\)" -AllowLabelAsValue
    $staffingAddition = Find-RowLabelValue -Cells $Cells -Row 3 -LabelPattern "人員配置体制加算" -AllowLabelAsValue
    $welfareSpecialist = Find-RowLabelValue -Cells $Cells -Row 3 -LabelPattern "福祉専門職員" -AllowLabelAsValue
    $behaviorTransition = Find-RowLabelValue -Cells $Cells -Row 4 -LabelPattern "地域移行体制特別加算" -AllowLabelAsValue
    $behaviorTrial = Find-RowLabelValue -Cells $Cells -Row 4 -LabelPattern "体験利用加算" -AllowLabelAsValue
    $peerSupport = Find-RowLabelValue -Cells $Cells -Row 4 -LabelPattern "ピアサポート" -AllowLabelAsValue
    $infectionControl = Find-RowLabelValue -Cells $Cells -Row 4 -LabelPattern "感染対策向上" -AllowLabelAsValue

    $labelRow = 0
    for ($row = 5; $row -le 12 -and $labelRow -eq 0; $row++) {
        for ($col = 1; $col -le 30; $col++) {
            $labelText = Normalize-LabelText (Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $col) + $row))
            if ($labelText -match "住居外利用") { $labelRow = $row; break }
        }
    }

    $columnPatterns = @(
        @{ Field = "住居外利用"; Pattern = "住居外利用" },
        @{ Field = "夜勤職員加配加算"; Pattern = "夜勤職員加配" },
        @{ Field = "入院時支援加算"; Pattern = "入院時支援" },
        @{ Field = "帰宅時支援加算"; Pattern = "帰宅時支援" },
        @{ Field = "医療連携体制加算"; Pattern = "医療連携" },
        @{ Field = "精神障害者地域移行特別加算"; Pattern = "精神障害者地域移行" },
        @{ Field = "地域生活移行個別支援特別加算"; Pattern = "地域生活移行" },
        @{ Field = "自立生活支援加算Ⅲ"; Pattern = "自立生活支援加算\(III\)" },
        @{ Field = "自立生活支援加算Ⅱ"; Pattern = "自立生活支援加算\(II\)" },
        @{ Field = "集中的支援加算Ⅱ"; Pattern = "集中的支援加算\(II\)" },
        @{ Field = "集中的支援加算Ⅰ"; Pattern = "集中的支援加算\(I\)" },
        @{ Field = "朝食"; Pattern = "^朝食" },
        @{ Field = "昼食"; Pattern = "^昼食" },
        @{ Field = "夕食"; Pattern = "^夕食" }
    )

    $columnMap = @{}
    $statusCol = 0
    if ($labelRow -gt 0) {
        for ($col = 1; $col -le 30; $col++) {
            $labelText = Normalize-LabelText (Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $col) + $labelRow))
            if ([string]::IsNullOrWhiteSpace($labelText)) { continue }
            if ($statusCol -eq 0 -and $labelText -match "サービス提供") { $statusCol = $col; continue }
            foreach ($columnPattern in $columnPatterns) {
                if (-not $columnMap.ContainsKey($columnPattern.Field) -and $labelText -match $columnPattern.Pattern) {
                    $columnMap[$columnPattern.Field] = $col
                    break
                }
            }
        }
    }

    $totalRow = 0
    if ($labelRow -gt 0) {
        for ($row = $labelRow + 1; $row -le $labelRow + 45; $row++) {
            # 旧版はA列しか見ていないため、合計ラベルが別列にある様式で合計行を見失っていました。
            for ($col = 1; $col -le 4; $col++) {
                $text = Normalize-LabelText (Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $col) + $row))
                if ($text -match "^合計") { $totalRow = $row; break }
            }
            if ($totalRow -gt 0) { break }
        }
    }

    if ($labelRow -eq 0 -or $totalRow -eq 0 -or $statusCol -eq 0) {
        Add-ReviewRow $FacilityName $MonthLabel $FileName ("利用者シート「{0}」のレイアウト（項目ラベル行・合計行）を特定できなかったため、実績詳細を抽出できませんでした。" -f $SheetName)
    }

    $hospitalDays = 0
    $overnightDays = 0
    $statusCounts = [ordered]@{}
    if ($labelRow -gt 0 -and $totalRow -gt 0 -and $statusCol -gt 0) {
        $statusLetter = Get-ColumnLetter $statusCol
        for ($row = $labelRow + 1; $row -lt $totalRow; $row++) {
            $status = To-Text (Get-CellValueFromMap -Cells $Cells -Address ($statusLetter + $row))
            if ([string]::IsNullOrWhiteSpace($status)) { continue }
            if (-not $statusCounts.Contains($status)) { $statusCounts[$status] = 0 }
            $statusCounts[$status] = $statusCounts[$status] + 1
            if ($status -match "入院") { $hospitalDays++ }
            if ($status -match "外泊") { $overnightDays++ }
        }
    }
    $statusBreakdown = (@($statusCounts.Keys | ForEach-Object { "{0}:{1}" -f $_, $statusCounts[$_] }) -join ", ")

    $basicDays = ""
    $totalsByField = @{}
    if ($totalRow -gt 0) {
        if ($statusCol -gt 0) {
            $basicDays = To-Text (Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $statusCol) + $totalRow))
        }
        foreach ($field in $columnMap.Keys) {
            $totalsByField[$field] = To-Text (Get-CellValueFromMap -Cells $Cells -Address ((Get-ColumnLetter $columnMap[$field]) + $totalRow))
        }
    }

    $properties = [ordered]@{
        "施設名" = $FacilityName
        "提供月" = $MonthLabel
        "種別" = $FileType
        "ファイル名" = $FileName
        "シート名" = $SheetName
        "受給者証番号" = $recipientNumber
        "氏名" = $userName
        "区分" = $categoryLevel
        "部屋番号" = $roomNumber
        "処遇改善加算" = $treatmentImprovement
        "重度障害者支援加算Ⅰ" = $severeSupport1
        "重度障害者支援加算Ⅱ" = $severeSupport2
        "人員配置体制加算" = $staffingAddition
        "福祉専門職員配置等加算" = $welfareSpecialist
        "強度行動障害者地域移行体制特別加算" = $behaviorTransition
        "強度行動障害者体験利用加算" = $behaviorTrial
        "ピアサポート実施加算" = $peerSupport
        "感染対策向上加算" = $infectionControl
        "基本算定日数" = $basicDays
    }

    foreach ($columnPattern in $columnPatterns) {
        $field = $columnPattern.Field
        $value = ""
        if ($totalsByField.ContainsKey($field)) { $value = $totalsByField[$field] }
        $properties[$field] = $value
    }

    $properties["入院日数"] = $hospitalDays.ToString()
    $properties["外泊日数"] = $overnightDays.ToString()
    $properties["状況内訳"] = $statusBreakdown

    return [PSCustomObject]$properties
}

function Get-UserRecordDetailHeaders {
    return @(
        "施設名", "提供月", "種別", "ファイル名", "シート名",
        "受給者証番号", "氏名", "区分", "部屋番号",
        "処遇改善加算", "重度障害者支援加算Ⅰ", "重度障害者支援加算Ⅱ", "人員配置体制加算", "福祉専門職員配置等加算",
        "強度行動障害者地域移行体制特別加算", "強度行動障害者体験利用加算", "ピアサポート実施加算", "感染対策向上加算",
        "基本算定日数", "住居外利用", "夜勤職員加配加算", "入院時支援加算", "帰宅時支援加算", "医療連携体制加算",
        "精神障害者地域移行特別加算", "地域生活移行個別支援特別加算",
        "自立生活支援加算Ⅲ", "自立生活支援加算Ⅱ", "集中的支援加算Ⅱ", "集中的支援加算Ⅰ",
        "朝食", "昼食", "夕食",
        "入院日数", "外泊日数", "状況内訳"
    )
}

# ==============================
# 集計
# ==============================
function Add-ErrorRow {
    param([string]$FacilityName, [string]$MonthLabel, [string]$FileName, [string]$ErrorMessage)
    $script:ErrorRows += [PSCustomObject]@{
        "施設名" = $FacilityName; "提供月" = $MonthLabel; "ファイル名" = $FileName; "エラー内容" = $ErrorMessage
    }
}

function Add-ReviewRow {
    param([string]$FacilityName, [string]$MonthLabel, [string]$FileName, [string]$Reason)
    $script:ReviewRows += [PSCustomObject]@{
        "施設名" = $FacilityName; "提供月" = $MonthLabel; "ファイル名" = $FileName; "確認理由" = $Reason
    }
}

function Add-SkippedRow {
    param([string]$FacilityName, [string]$MonthLabel, [string]$FileName, [string]$Reason)
    $script:SkippedRows += [PSCustomObject]@{
        "施設名" = $FacilityName; "提供月" = $MonthLabel; "ファイル名" = $FileName; "対象外理由" = $Reason
    }
}

function Get-CategoryColumnLabel {
    param($Category)

    $text = To-Text $Category
    if ([string]::IsNullOrWhiteSpace($text)) { return $BlankCategoryLabel }
    return $text
}

# 入居（1F/2F/3F/その他）の施設単位合計行を作ります。
# 旧版の「1F+2F合計」は施設名の正規化が階判定と食い違い、岩沼ⅰ/ⅱ・美田園北/南・白浜町Ⅰ/Ⅱ では
# 実際には合算されていませんでした。ここでは統一済みの施設名で必ず合算します。
function Add-AdmissionTotalRows {
    param([object[]]$Rows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return @($Rows) }

    $resultRows = @($Rows)
    $admissionRows = @($Rows | Where-Object { $AdmissionTypes -contains $_."種別" })
    if ($admissionRows.Count -eq 0) { return @($resultRows) }

    # 施設・月ごとのK5合計（同じファイルのK5を二重に足さないよう、ファイル単位で1回だけ加算）
    $totalByFacilityMonth = @{}
    $seenFiles = @{}
    foreach ($row in $admissionRows) {
        $facilityMonthKey = "{0}`t{1}" -f $row."施設名", $row."提供月"
        $fileKey = "{0}`t{1}" -f $facilityMonthKey, $row."ファイル名"
        if ($seenFiles.ContainsKey($fileKey)) { continue }
        $seenFiles[$fileKey] = $true
        if (-not $totalByFacilityMonth.ContainsKey($facilityMonthKey)) { $totalByFacilityMonth[$facilityMonthKey] = 0.0 }
        $totalByFacilityMonth[$facilityMonthKey] += Convert-ToNumberOrZero $row."利用回数合計"
    }

    $categoryGroups = $admissionRows | Group-Object { "{0}`t{1}`t{2}" -f $_."施設名", $_."提供月", (Get-CategoryColumnLabel $_."区分") }
    foreach ($categoryGroup in $categoryGroups) {
        $firstRow = $categoryGroup.Group | Select-Object -First 1
        $usageTotal = 0.0
        foreach ($row in $categoryGroup.Group) { $usageTotal += Convert-ToNumberOrZero $row."利用回数" }

        $facilityMonthKey = "{0}`t{1}" -f $firstRow."施設名", $firstRow."提供月"
        $allUsageTotal = 0.0
        if ($totalByFacilityMonth.ContainsKey($facilityMonthKey)) { $allUsageTotal = $totalByFacilityMonth[$facilityMonthKey] }

        $resultRows += [PSCustomObject]@{
            "施設名" = $firstRow."施設名"
            "棟" = ""
            "フォルダ施設名" = $firstRow."フォルダ施設名"
            "提供月" = $firstRow."提供月"
            "種別" = "入居合計"
            "ファイル名" = "（入居合計）"
            "区分" = $firstRow."区分"
            "利用回数" = Format-NumberForCsv $usageTotal
            "利用回数合計" = Format-NumberForCsv $allUsageTotal
        }
    }

    return @(
        $resultRows | Sort-Object `
            @{ Expression = { $_."施設名" } },
            @{ Expression = { $_."提供月" } },
            @{ Expression = { Get-CategoryTypeSortKey $_."種別" } },
            @{ Expression = { Get-CategorySortKey (Get-CategoryColumnLabel $_."区分") } }
    )
}

# シート「施設別合計」。元データ（入居合計行を含まない生の区分別行）から直接作ります。
function Get-CategoryRowsForSummaryWorkbook {
    param([object[]]$Rows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return @() }

    $sourceRows = @($Rows | Where-Object { $_."種別" -ne "入居合計" })
    if ($sourceRows.Count -eq 0) { return @() }

    $groups = $sourceRows | Group-Object { "{0}`t{1}`t{2}" -f $_."施設名", $_."提供月", (Get-CategoryColumnLabel $_."区分") }

    $summaryRows = @()
    foreach ($group in $groups) {
        $firstRow = $group.Group | Select-Object -First 1
        $admission12 = 0.0
        $admission3 = 0.0
        $admissionOther = 0.0
        $shortStay = 0.0

        foreach ($row in $group.Group) {
            $value = Convert-ToNumberOrZero $row."利用回数"
            switch ($row."種別") {
                "入居1F"     { $admission12 += $value }
                "入居2F"     { $admission12 += $value }
                "入居3F"     { $admission3 += $value }
                "入居その他" { $admissionOther += $value }
                "短期"       { $shortStay += $value }
            }
        }

        $summaryRows += [PSCustomObject]@{
            "施設名" = $firstRow."施設名"
            "提供月" = $firstRow."提供月"
            "区分" = Get-CategoryColumnLabel $firstRow."区分"
            "入居1F+2F" = Format-NumberForCsv $admission12
            "入居3F" = Format-NumberForCsv $admission3
            "入居その他" = Format-NumberForCsv $admissionOther
            "短期" = Format-NumberForCsv $shortStay
            "合計" = Format-NumberForCsv ($admission12 + $admission3 + $admissionOther + $shortStay)
        }
    }

    return @(
        $summaryRows | Sort-Object `
            @{ Expression = { $_."施設名" } },
            @{ Expression = { $_."提供月" } },
            @{ Expression = { Get-CategorySortKey $_."区分" } }
    )
}

# 施設別・区分別のマトリクス。区分が空欄の行も列として残すので、合計が元データと一致します。
function Get-FacilityCategoryMatrixRows {
    param(
        [object[]]$Rows,
        [string[]]$IncludeTypes
    )

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return @() }

    $sourceRows = @($Rows | Where-Object { $_."種別" -ne "入居合計" -and $IncludeTypes -contains $_."種別" })
    if ($sourceRows.Count -eq 0) { return @() }

    $categories = @($sourceRows | ForEach-Object { Get-CategoryColumnLabel $_."区分" } | Sort-Object -Unique)
    $categories = @($categories | Sort-Object { Get-CategorySortKey $_ })

    $totals = @{}
    $facilityMonthMap = @{}
    foreach ($row in $sourceRows) {
        $facilityName = To-Text $row."施設名"
        $monthLabel = To-Text $row."提供月"
        $facilityMonthKey = "{0}`t{1}" -f $facilityName, $monthLabel
        if (-not $facilityMonthMap.ContainsKey($facilityMonthKey)) {
            $facilityMonthMap[$facilityMonthKey] = [PSCustomObject]@{ Facility = $facilityName; Month = $monthLabel }
        }

        $category = Get-CategoryColumnLabel $row."区分"
        $categoryKey = "{0}`t{1}" -f $facilityMonthKey, $category
        if (-not $totals.ContainsKey($categoryKey)) { $totals[$categoryKey] = 0.0 }
        $totals[$categoryKey] += Convert-ToNumberOrZero $row."利用回数"
    }

    $matrixRows = @()
    $grandTotalByCategory = @{}
    $grandTotalAll = 0.0

    foreach ($facilityMonth in @($facilityMonthMap.Values | Sort-Object Facility, Month)) {
        $properties = [ordered]@{
            "施設名" = $facilityMonth.Facility
            "提供月" = $facilityMonth.Month
        }
        $rowTotal = 0.0
        foreach ($category in $categories) {
            $categoryKey = "{0}`t{1}`t{2}" -f $facilityMonth.Facility, $facilityMonth.Month, $category
            $value = 0.0
            if ($totals.ContainsKey($categoryKey)) { $value = $totals[$categoryKey] }
            $properties[(Get-SafeCategoryColumnName $category)] = Format-NumberForCsv $value
            $rowTotal += $value

            if (-not $grandTotalByCategory.ContainsKey($category)) { $grandTotalByCategory[$category] = 0.0 }
            $grandTotalByCategory[$category] += $value
        }
        $properties["合計"] = Format-NumberForCsv $rowTotal
        $grandTotalAll += $rowTotal

        $matrixRows += [PSCustomObject]$properties
    }

    $totalProperties = [ordered]@{ "施設名" = "合計"; "提供月" = "" }
    foreach ($category in $categories) {
        $value = 0.0
        if ($grandTotalByCategory.ContainsKey($category)) { $value = $grandTotalByCategory[$category] }
        $totalProperties[(Get-SafeCategoryColumnName $category)] = Format-NumberForCsv $value
    }
    $totalProperties["合計"] = Format-NumberForCsv $grandTotalAll
    $matrixRows += [PSCustomObject]$totalProperties

    return @($matrixRows)
}

function Get-MatrixHeaders {
    param([object[]]$Rows)

    if ($null -ne $Rows -and $Rows.Count -gt 0) {
        return @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    }
    return @("施設名", "提供月", "合計")
}

function Get-RowsTotal {
    param(
        [object[]]$Rows,
        [string]$ColumnName
    )

    $total = 0.0
    if ($null -eq $Rows) { return $total }
    foreach ($row in $Rows) {
        if ($null -eq $row.PSObject.Properties[$ColumnName]) { continue }
        $total += Convert-ToNumberOrZero $row.$ColumnName
    }
    return $total
}

# マトリクスの最終行は全施設合計なので、検算では除外します。
function Get-MatrixTotal {
    param([object[]]$Rows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return 0.0 }
    $bodyRows = @($Rows | Where-Object { $_."施設名" -ne "合計" })
    return Get-RowsTotal -Rows $bodyRows -ColumnName "合計"
}

# ==============================
# Excel出力
# ==============================
# Excelが文字列を勝手に日付・数値へ変換するのを防ぐため、文字列として扱う列を指定します。
# （旧版は部屋番号「1-2」等が日付に化けていました）
$TextColumnNames = @(
    "提供月", "受給者証番号", "部屋番号", "ファイル名", "シート名", "状況内訳", "棟", "フォルダ施設名"
)

function Write-RowsToWorksheet {
    param(
        $Worksheet,
        [object[]]$Rows,
        [string[]]$Headers,
        [int]$StartRow = 1,
        [string[]]$ExtraTextColumns = @(),
        [switch]$NoFreeze
    )

    $textColumns = @($TextColumnNames + $ExtraTextColumns)

    for ($col = 1; $col -le $Headers.Count; $col++) {
        if ($textColumns -contains $Headers[$col - 1]) {
            $column = $Worksheet.Columns.Item($col)
            $column.NumberFormat = "@"
            Release-ComObject $column
        }
        $Worksheet.Cells.Item($StartRow, $col).Value2 = $Headers[$col - 1]
    }

    $rowIndex = $StartRow + 1
    foreach ($row in $Rows) {
        for ($col = 1; $col -le $Headers.Count; $col++) {
            $header = $Headers[$col - 1]
            $value = ""
            if ($null -ne $row.PSObject.Properties[$header]) { $value = To-Text $row.$header }
            $Worksheet.Cells.Item($rowIndex, $col).Value2 = $value
        }
        $rowIndex++
    }

    $usedRange = $Worksheet.UsedRange
    $usedRange.Font.Name = "Meiryo UI"
    $usedRange.Font.Size = 10
    $headerRange = $Worksheet.Range($Worksheet.Cells.Item($StartRow, 1), $Worksheet.Cells.Item($StartRow, $Headers.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 14277081
    $usedRange.Columns.AutoFit() | Out-Null

    # 固定行は「アクティブなウィンドウ」に効くため、対象シートを必ずアクティブにしてから設定します。
    # （サマリーのように1シートへ複数の表を書く場合は $NoFreeze で抑止します）
    if (-not $NoFreeze) {
        $Worksheet.Activate()
        $Worksheet.Application.ActiveWindow.FreezePanes = $false
        $Worksheet.Application.ActiveWindow.SplitRow = $StartRow
        $Worksheet.Application.ActiveWindow.SplitColumn = 0
        $Worksheet.Application.ActiveWindow.FreezePanes = $true
    }

    Release-ComObject $headerRange
    Release-ComObject $usedRange
}

function Write-SummarySheet {
    param(
        $Worksheet,
        [PSCustomObject]$SummaryInfo,
        [object[]]$ReconciliationRows,
        [object[]]$FacilityGrandTotals
    )

    $infoColumn = $Worksheet.Columns.Item(2)
    $infoColumn.NumberFormat = "@"
    Release-ComObject $infoColumn

    $Worksheet.Cells.Item(1, 1).Value2 = "実行サマリー"
    $Worksheet.Cells.Item(1, 1).Font.Bold = $true

    $row = 3
    foreach ($property in $SummaryInfo.PSObject.Properties) {
        $Worksheet.Cells.Item($row, 1).Value2 = $property.Name
        $Worksheet.Cells.Item($row, 1).Font.Bold = $true
        $Worksheet.Cells.Item($row, 2).Value2 = To-Text $property.Value
        $row++
    }

    $reconStartRow = $row + 1
    $Worksheet.Cells.Item($reconStartRow, 1).Value2 = "検算（元データの利用回数と各シートの合計が一致するか）"
    $Worksheet.Cells.Item($reconStartRow, 1).Font.Bold = $true
    Write-RowsToWorksheet -Worksheet $Worksheet -Rows $ReconciliationRows `
        -Headers @("検算対象", "元データ合計", "シート合計", "差分", "判定") -StartRow ($reconStartRow + 1) -NoFreeze

    $tableStartRow = $reconStartRow + $ReconciliationRows.Count + 3
    $Worksheet.Cells.Item($tableStartRow, 1).Value2 = "施設別合計利用回数"
    $Worksheet.Cells.Item($tableStartRow, 1).Font.Bold = $true

    Write-RowsToWorksheet -Worksheet $Worksheet -Rows $FacilityGrandTotals `
        -Headers @("施設名", "提供月", "合計利用回数") -StartRow ($tableStartRow + 1) -NoFreeze
}

function Write-ResultWorkbook {
    param(
        [PSCustomObject]$SummaryInfo,
        [object[]]$ReconciliationRows,
        [object[]]$FacilityGrandTotals,
        [object[]]$AdmissionMatrixRows,
        [object[]]$ShortStayMatrixRows,
        [object[]]$CombinedMatrixRows,
        [object[]]$CategorySummaryRows,
        [object[]]$CategoryDetailRows,
        [object[]]$UserRows,
        [object[]]$UserDetailRows,
        [object[]]$ErrorRows,
        [object[]]$ReviewRows,
        [object[]]$SkippedRows,
        [string]$Path
    )

    $workbookExcel = $null
    $workbook = $null
    $sheets = @{}
    $originalSheetCount = $null

    try {
        Write-Host ("結果Excelを作成しています: {0}" -f $Path)
        $workbookExcel = New-Object -ComObject Excel.Application
        $workbookExcel.Visible = $false
        $workbookExcel.DisplayAlerts = $false

        # SheetsInNewWorkbook はExcelアプリ全体の永続設定なので、必ず元に戻します。
        $originalSheetCount = $workbookExcel.SheetsInNewWorkbook
        $workbookExcel.SheetsInNewWorkbook = 11

        $workbook = $workbookExcel.Workbooks.Add()

        $sheetNames = @(
            "サマリー",
            "施設別・区分別合計（入居）",
            "施設別・区分別合計（短期）",
            "施設別・区分別合計（入居＋短期）",
            "施設別合計",
            "区分別_利用回数一覧",
            "利用者別_延べ日数一覧",
            "利用者別_実績詳細",
            "エラー一覧",
            "確認対象一覧",
            "対象外ファイル一覧"
        )
        for ($i = 0; $i -lt $sheetNames.Count; $i++) {
            $sheet = $workbook.Worksheets.Item($i + 1)
            $sheet.Name = $sheetNames[$i]
            $sheets[$sheetNames[$i]] = $sheet
        }

        Write-SummarySheet -Worksheet $sheets["サマリー"] -SummaryInfo $SummaryInfo `
            -ReconciliationRows $ReconciliationRows -FacilityGrandTotals $FacilityGrandTotals
        Write-RowsToWorksheet -Worksheet $sheets["施設別・区分別合計（入居）"] -Rows $AdmissionMatrixRows -Headers (Get-MatrixHeaders -Rows $AdmissionMatrixRows)
        Write-RowsToWorksheet -Worksheet $sheets["施設別・区分別合計（短期）"] -Rows $ShortStayMatrixRows -Headers (Get-MatrixHeaders -Rows $ShortStayMatrixRows)
        Write-RowsToWorksheet -Worksheet $sheets["施設別・区分別合計（入居＋短期）"] -Rows $CombinedMatrixRows -Headers (Get-MatrixHeaders -Rows $CombinedMatrixRows)
        Write-RowsToWorksheet -Worksheet $sheets["施設別合計"] -Rows $CategorySummaryRows `
            -Headers @("施設名", "提供月", "区分", "入居1F+2F", "入居3F", "入居その他", "短期", "合計") -ExtraTextColumns @("区分")
        Write-RowsToWorksheet -Worksheet $sheets["区分別_利用回数一覧"] -Rows $CategoryDetailRows `
            -Headers @("施設名", "棟", "フォルダ施設名", "提供月", "種別", "ファイル名", "区分", "利用回数", "利用回数合計") -ExtraTextColumns @("区分")
        Write-RowsToWorksheet -Worksheet $sheets["利用者別_延べ日数一覧"] -Rows $UserRows `
            -Headers @("施設名", "棟", "提供月", "種別", "ファイル名", "利用者名", "区分", "延べ日数") -ExtraTextColumns @("区分")
        Write-RowsToWorksheet -Worksheet $sheets["利用者別_実績詳細"] -Rows $UserDetailRows -Headers (Get-UserRecordDetailHeaders) -ExtraTextColumns @("区分")
        Write-RowsToWorksheet -Worksheet $sheets["エラー一覧"] -Rows $ErrorRows -Headers @("施設名", "提供月", "ファイル名", "エラー内容")
        Write-RowsToWorksheet -Worksheet $sheets["確認対象一覧"] -Rows $ReviewRows -Headers @("施設名", "提供月", "ファイル名", "確認理由")
        Write-RowsToWorksheet -Worksheet $sheets["対象外ファイル一覧"] -Rows $SkippedRows -Headers @("施設名", "提供月", "ファイル名", "対象外理由")

        $sheets["サマリー"].Activate()

        try {
            if (Test-Path -LiteralPath $Path) {
                Remove-Item -LiteralPath $Path -Force
            }
            $workbook.SaveAs($Path, 51)
            return $Path
        }
        catch {
            $directory = Split-Path -Parent $Path
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $fallbackPath = Join-Path $directory ("{0}_{1}.xlsx" -f $baseName, $timestamp)
            $workbook.SaveAs($fallbackPath, 51)
            Write-Host ("出力ファイルが開かれていたため、別名で保存しました: {0}" -f $fallbackPath) -ForegroundColor Yellow
            return $fallbackPath
        }
    }
    finally {
        if ($null -ne $workbookExcel -and $null -ne $originalSheetCount) {
            try { $workbookExcel.SheetsInNewWorkbook = $originalSheetCount } catch { }
        }
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $workbookExcel) { $workbookExcel.Quit() }
        foreach ($sheet in $sheets.Values) { Release-ComObject $sheet }
        Release-ComObject $workbook
        Release-ComObject $workbookExcel
    }
}

# ==============================
# 対象月・対象施設の選択
# ==============================
function Select-TargetMonthKey {
    param([string[]]$AvailableMonthKeys)

    $uniqueMonthKeys = @($AvailableMonthKeys | Sort-Object -Unique)
    $latestMonthKey = $uniqueMonthKeys | Sort-Object -Descending | Select-Object -First 1
    $latestYear = ""
    $latestMonth = ""

    if ($latestMonthKey -match "^(\d{4})-(\d{2})$") {
        $latestYear = $Matches[1]
        $latestMonth = [string][int]$Matches[2]
    }

    Write-Host ""
    Write-Host "見つかった提供月:"
    foreach ($monthKey in $uniqueMonthKeys) { Write-Host ("  - " + (Format-MonthLabel $monthKey)) }
    Write-Host ""

    while ($true) {
        $yearText = Read-Host "対象年を入力してください（例: 2026。空欄なら $latestYear）"
        $monthText = Read-Host "対象月を入力してください（例: 6。空欄なら $latestMonth）"

        if ([string]::IsNullOrWhiteSpace($yearText)) { $yearText = $latestYear }
        if ([string]::IsNullOrWhiteSpace($monthText)) { $monthText = $latestMonth }

        $yearText = (Normalize-Text $yearText).Trim()
        $monthText = (Normalize-Text $monthText).Trim()

        $year = 0
        $month = 0
        $isValid = [int]::TryParse($yearText, [ref]$year) -and [int]::TryParse($monthText, [ref]$month) -and $year -ge 2000 -and $month -ge 1 -and $month -le 12

        if (-not $isValid) {
            Write-Host "年または月の入力が正しくありません。例: 年=2026、月=6" -ForegroundColor Yellow
            Write-Host ""
            continue
        }

        $selectedMonthKey = "{0:D4}-{1:D2}" -f $year, $month
        if ($uniqueMonthKeys -contains $selectedMonthKey) { return $selectedMonthKey }

        Write-Host ("指定した年月の月フォルダが見つかりません: " + (Format-MonthLabel $selectedMonthKey)) -ForegroundColor Yellow
        Write-Host "もう一度入力してください。"
        Write-Host ""
    }
}

function Get-FacilityKeywords {
    param([string]$KeywordText)

    return @(
        $KeywordText -split "[,、，]" |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Get-MatchedFacilityNames {
    param(
        [string[]]$AvailableFacilityNames,
        [string[]]$Keywords
    )

    return @(
        $AvailableFacilityNames | Where-Object {
            $facilityName = $_
            $isMatch = $false
            foreach ($keyword in $Keywords) {
                if ($facilityName -like "*$keyword*") { $isMatch = $true; break }
            }
            $isMatch
        }
    )
}

function Select-TargetFacility {
    param([string[]]$AvailableFacilityNames)

    $uniqueFacilityNames = @($AvailableFacilityNames | Sort-Object -Unique)

    Write-Host ""
    Write-Host "見つかった施設:"
    for ($i = 0; $i -lt $uniqueFacilityNames.Count; $i++) {
        Write-Host ("  {0,3}: {1}" -f ($i + 1), $uniqueFacilityNames[$i])
    }
    Write-Host ""

    while ($true) {
        $inputText = Read-Host "対象施設の番号を入力してください（複数はカンマ区切り。全角数字も可。施設名の一部でも指定できます。空欄なら全施設）"

        if ([string]::IsNullOrWhiteSpace($inputText)) { return @() }

        $normalizedInput = Normalize-Text $inputText
        $tokens = @($normalizedInput -split "[,、，\s]+" | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        $allTokensAreNumbers = ($tokens.Count -gt 0)
        foreach ($token in $tokens) {
            $parsedNumber = 0
            if (-not [int]::TryParse($token, [ref]$parsedNumber)) { $allTokensAreNumbers = $false; break }
        }

        if ($allTokensAreNumbers) {
            $selectedNames = @()
            $invalidNumbers = @()
            foreach ($token in $tokens) {
                $number = [int]$token
                if ($number -ge 1 -and $number -le $uniqueFacilityNames.Count) {
                    $facilityName = $uniqueFacilityNames[$number - 1]
                    if ($selectedNames -notcontains $facilityName) { $selectedNames += $facilityName }
                } else {
                    $invalidNumbers += $token
                }
            }

            if ($invalidNumbers.Count -gt 0) {
                Write-Host ("番号が範囲外です: {0}（1～{1} の番号で指定してください）" -f ($invalidNumbers -join ", "), $uniqueFacilityNames.Count) -ForegroundColor Yellow
                Write-Host ""
                continue
            }

            Write-Host ""
            Write-Host ("選択した施設: {0} 件" -f $selectedNames.Count)
            foreach ($selectedName in $selectedNames) { Write-Host ("  - " + $selectedName) }
            Write-Host ""
            return $selectedNames
        }

        $keywords = @(Get-FacilityKeywords -KeywordText $inputText)
        $matchedFacilityNames = @(Get-MatchedFacilityNames -AvailableFacilityNames $uniqueFacilityNames -Keywords $keywords)

        if ($matchedFacilityNames.Count -eq 0) {
            Write-Host ("指定した施設名「{0}」に一致する施設が見つかりません。もう一度入力してください。" -f $inputText) -ForegroundColor Yellow
            Write-Host ""
            continue
        }

        Write-Host ""
        Write-Host ("一致した施設: {0} 件" -f $matchedFacilityNames.Count)
        foreach ($matchedName in $matchedFacilityNames) { Write-Host ("  - " + $matchedName) }
        Write-Host ""

        if ($matchedFacilityNames.Count -gt 1) {
            $continueAnswer = Read-Host "複数の施設が一致しました。すべて対象にして続行しますか？ (Y/N)"
            if ($continueAnswer -notmatch "^(y|yes|はい)$") {
                Write-Host "もう一度入力してください。"
                Write-Host ""
                continue
            }
        }

        return $matchedFacilityNames
    }
}

# ==============================
# メイン処理
# ==============================
if (-not (Test-Path -LiteralPath $RootPath)) {
    throw "対象フォルダが見つかりません: $RootPath"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$UserRows = @()
$UserDetailRows = @()
$CategoryRows = @()
$ErrorRows = @()
$ReviewRows = @()
$SkippedRows = @()

Write-Host "対象ルート: $RootPath"
Write-Host "月フォルダを検索しています..."

try {
    $allMonthFolders = @(Get-ChildItem -LiteralPath $RootPath -Directory -Recurse |
        ForEach-Object { Get-MonthFolderInfo $_ } |
        Where-Object { $null -ne $_ })

    if ($allMonthFolders.Count -eq 0) {
        throw "月フォルダが見つかりません。フォルダ名に「2026年6月」「2026年6月分」、または「2026年\6月」のような年月が含まれているか確認してください。"
    }

    if ([string]::IsNullOrWhiteSpace($TargetMonth)) {
        $requestedMonthKey = Select-TargetMonthKey -AvailableMonthKeys @($allMonthFolders | Select-Object -ExpandProperty MonthKey)
    } else {
        $requestedMonthKey = Get-MonthKeyFromText $TargetMonth
        if ($null -eq $requestedMonthKey) { throw "TargetMonthの指定を年月として読み取れません: $TargetMonth" }
    }

    $monthFolders = @(
        $allMonthFolders |
            Where-Object { $_.MonthKey -eq $requestedMonthKey } |
            ForEach-Object {
                $facilityNameForFolder = Get-FacilityName -Root $RootPath -MonthFolderPath $_.Path
                [PSCustomObject]@{
                    Path = $_.Path
                    Name = $_.Name
                    MonthKey = $_.MonthKey
                    LastWriteTime = $_.LastWriteTime
                    ContentLastWriteTime = Get-FolderContentLastWriteTime $_.Path
                    Priority = $_.Priority
                    FacilityName = $facilityNameForFolder
                }
            } |
            Group-Object { "{0}`t{1}" -f $_.FacilityName, $_.MonthKey } |
            ForEach-Object {
                $group = $_.Group
                if ($group.Count -gt 1) {
                    $bestPriority = ($group | Measure-Object -Property Priority -Minimum).Minimum
                    $topTierFolders = @($group | Where-Object { $_.Priority -eq $bestPriority })
                    if ($topTierFolders.Count -gt 1) {
                        $folderList = ($topTierFolders | ForEach-Object { "{0}（フォルダ内Excel更新日時: {1}）" -f $_.Path, $_.ContentLastWriteTime.ToString("yyyy/MM/dd HH:mm:ss") }) -join " / "
                        Add-ReviewRow $group[0].FacilityName (Format-MonthLabel $group[0].MonthKey) "" ("同一施設・同一月のフォルダが複数見つかり、「請求用」「作成用」等の名称だけでは一意に絞り込めませんでした。フォルダ内のExcel更新日時が最も新しいものを自動選択しましたが、内容を確認してください: {0}" -f $folderList)
                    }
                }
                $group | Sort-Object Priority, @{Expression = "ContentLastWriteTime"; Descending = $true} | Select-Object -First 1
            }
    )

    if ($monthFolders.Count -eq 0) {
        throw "対象月フォルダが見つかりません。TargetMonthの指定、またはフォルダ名を確認してください。"
    }

    $availableFacilityNames = @($monthFolders | Select-Object -ExpandProperty FacilityName -Unique)

    if ([string]::IsNullOrWhiteSpace($TargetFacility)) {
        $selectedFacilityNames = @(Select-TargetFacility -AvailableFacilityNames $availableFacilityNames)
    } else {
        $targetFacilityKeywords = @(Get-FacilityKeywords -KeywordText $TargetFacility)
        $selectedFacilityNames = @(Get-MatchedFacilityNames -AvailableFacilityNames $availableFacilityNames -Keywords $targetFacilityKeywords)

        if ($selectedFacilityNames.Count -eq 0) {
            throw "指定した施設名「$TargetFacility」に一致する施設が見つかりません。TargetFacilityの指定を確認してください。"
        }

        Write-Host ""
        Write-Host ("施設名「{0}」に一致した施設: {1} 件" -f $TargetFacility, $selectedFacilityNames.Count)
        foreach ($matchedName in $selectedFacilityNames) { Write-Host ("  - " + $matchedName) }
        Write-Host ""

        if ($selectedFacilityNames.Count -gt 1) {
            $continueAnswer = Read-Host "複数の施設が一致しました。すべて対象にして続行しますか？ (Y/N)"
            if ($continueAnswer -notmatch "^(y|yes|はい)$") {
                throw "処理を中止しました。TargetFacilityの指定を見直してください。"
            }
        }
    }

    if ($selectedFacilityNames.Count -gt 0) {
        $monthFolders = @($monthFolders | Where-Object { $selectedFacilityNames -contains $_.FacilityName })
    }
}
catch {
    Write-Host ""
    Write-Host "対象施設・月フォルダの検索中にエラーが発生したため、処理を中止しました。どの施設も処理できていません。" -ForegroundColor Red
    Write-Host ("エラー内容: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host ""
    Write-Host "Dropboxの同期状態（フォルダが「オンラインのみ」になっていないか）、対象フォルダへのアクセス権限、TargetFacilityの指定を確認してから再実行してください。" -ForegroundColor Yellow
    exit 1
}

$targetMonthLabels = $monthFolders | Select-Object -ExpandProperty MonthKey -Unique | ForEach-Object { Format-MonthLabel $_ }
Write-Host ("抽出対象月: " + ($targetMonthLabels -join ", "))
Write-Host ("対象月フォルダ数: " + $monthFolders.Count)

$excel = $null

try {
    $monthFolderIndex = 0
    foreach ($monthFolder in $monthFolders) {
        $monthFolderIndex++
        $folderFacilityName = Get-FacilityName -Root $RootPath -MonthFolderPath $monthFolder.Path
        $monthLabel = Format-MonthLabel $monthFolder.MonthKey
        Write-Host ("施設処理中: {0}/{1} {2}" -f $monthFolderIndex, $monthFolders.Count, $folderFacilityName)

        try {

        $daysInMonth = 31
        if ($monthFolder.MonthKey -match "^(\d{4})-(\d{2})$") {
            $daysInMonth = [DateTime]::DaysInMonth([int]$Matches[1], [int]$Matches[2])
        }

        if ([string]::IsNullOrWhiteSpace($folderFacilityName) -or $null -ne (Get-MonthKeyFromText $folderFacilityName)) {
            Add-ErrorRow $folderFacilityName $monthLabel "" ("施設名の自動判定結果が不審です（値: '{0}'）。「ルート直下1階層目＝施設名」という前提とフォルダ構成が異なる可能性があります。パス: {1}" -f $folderFacilityName, $monthFolder.Path)
        }

        $allExcelFilesInFolder = @(Get-ChildItem -LiteralPath $monthFolder.Path -File -Recurse |
            Where-Object {
                $_.Extension -in @(".xlsx", ".xlsm", ".xls") -and
                $_.Name -notlike '~$*' -and
                $_.Name -notlike "*.tmp"
            })

        $allExcelFiles = @()
        foreach ($candidateFile in $allExcelFilesInFolder) {
            if (Is-RecordCandidateExcelFile $candidateFile.Name) {
                $allExcelFiles += $candidateFile
            } elseif (Is-ExcludedExcelFile $candidateFile.Name) {
                Add-SkippedRow $folderFacilityName $monthLabel $candidateFile.Name "除外キーワード（受給者証/職員名簿/家賃/医療連携/届出書/請求金額/業務日誌/注意点/変更点/ブランク/様式。ただし「新様式」の実績記録票は除外しません）に一致したため対象外としました。"
            } else {
                Add-SkippedRow $folderFacilityName $monthLabel $candidateFile.Name "実績記録系のキーワード（実績記録/実績記録票/日中サービス支援/介護包括/短期）に一致しなかったため対象外としました。集計漏れの可能性があるためファイル名を確認してください。"
            }
        }

        # ファイルごとに施設・棟・種別を確定します。
        $typedFiles = @()
        foreach ($candidateFile in $allExcelFiles) {
            $identity = Get-FacilityIdentity -FileName $candidateFile.Name -FolderFacilityName $folderFacilityName
            $candidateType = Get-FileType -FileName $candidateFile.Name -Identity $identity

            if (-not $identity.Matched) {
                Add-ReviewRow $folderFacilityName $monthLabel $candidateFile.Name ("施設定義テーブルに一致しないため、ファイル名から作った「{0}」を施設名として集計しました。$FacilityRules に施設を追加すると安定します。" -f $identity.Facility)
            }
            if (Test-IsFloorUndetermined -FileName $candidateFile.Name -Identity $identity) {
                Add-ReviewRow $folderFacilityName $monthLabel $candidateFile.Name ("階（1F/2F/3F）を判定できなかったため、種別「{0}」として集計しました（集計対象からは外していません）。" -f $candidateType)
            }
            if ($identity.Facility -like "*(棟共通)") {
                Add-ReviewRow $folderFacilityName $monthLabel $candidateFile.Name ("棟を分けて集計する施設ですが、ファイル名から棟（ⅰ/ⅱ等）を特定できませんでした。どの棟にも寄せず「{0}」として独立集計しています。" -f $identity.Facility)
            }

            $typedFiles += [PSCustomObject]@{
                File = $candidateFile
                Type = $candidateType
                Identity = $identity
                LastWriteTime = $candidateFile.LastWriteTime
            }
        }

        # 同一「棟＋種別」の重複のみを1件に絞ります。
        # 旧版は種別だけで絞っていたため、萩丘ⅰ1F と 萩丘ⅱ1F のような別棟の同種別ファイルが片方消えていました。
        $excelFiles = @(
            $typedFiles |
                Group-Object { "{0}`t{1}" -f $_.Identity.Unit, $_.Type } |
                ForEach-Object {
                    $groupKeyName = $_.Group[0].Identity.Unit + " / " + $_.Group[0].Type
                    $selectedFile = $null
                    foreach ($candidate in ($_.Group | Sort-Object LastWriteTime -Descending)) {
                        if ($null -ne $selectedFile) {
                            Add-ReviewRow $folderFacilityName $monthLabel $candidate.File.Name ("同一棟・同一種別（{0}）のファイルが複数あるため、更新日時が新しい方を採用しました。このファイルは集計対象外です。採用ファイル: {1}（更新日時 {2}）" -f $groupKeyName, $selectedFile.File.Name, $selectedFile.File.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss"))
                            continue
                        }
                        if (Test-CanReadExcelFile $candidate.File.FullName) { $selectedFile = $candidate; continue }
                        Add-ErrorRow $folderFacilityName $monthLabel $candidate.File.Name "ファイルが使用中で読めないため、次に新しい同一棟・同一種別ファイルを確認しました。"
                    }
                    if ($null -ne $selectedFile) { $selectedFile }
                }
        )

        foreach ($fileInfo in $excelFiles) {
            $file = $fileInfo.File
            $fileType = $fileInfo.Type
            $identity = $fileInfo.Identity
            $facilityName = $identity.Facility

            $workbook = $null
            $sheet = $null

            try {
                Write-Host ("読込中: {0} / {1} / {2} / {3} / 更新日時 {4}" -f $facilityName, $monthLabel, $fileType, $file.Name, $file.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss"))

                if ($file.Extension -in @(".xlsx", ".xlsm")) {
                    # ブックは1回だけ開いて全シートを読み、その中から対象シートを選びます。
                    $workbookSheets = @(Read-OpenXmlWorkbookSheets -Path $file.FullName)
                    $sheetData = Select-TargetSheet -Sheets $workbookSheets -SheetNames $TargetSheetNames

                    if ($null -eq $sheetData) {
                        Add-ErrorRow $facilityName $monthLabel $file.Name "対象シート「$($TargetSheetNames -join ' / ')」が見つかりません。区分別・利用者別の集計をスキップしました。"
                    } else {
                        $cells = $sheetData.Cells
                        $cellErrors = $sheetData.CellErrors

                        $userColumns = Find-UserTableColumnsFromMap -Cells $cells
                        $scanStartRow = 1
                        if ($null -eq $userColumns) {
                            $nameColLetter = "B"; $categoryColLetter = "C"; $daysColLetter = "D"
                            Add-ReviewRow $facilityName $monthLabel $file.Name "利用者一覧のヘッダー（利用者名/区分/延べ日数）を自動検出できなかったため、既定のB/C/D列を使用しました。レイアウトを確認してください。"
                        } else {
                            $nameColLetter = Get-ColumnLetter $userColumns.NameCol
                            $categoryColLetter = Get-ColumnLetter $userColumns.CategoryCol
                            $daysColLetter = Get-ColumnLetter $userColumns.DaysCol
                            # ヘッダー行より上（タイトル行など）は読まない
                            $scanStartRow = $userColumns.HeaderRow + 1
                        }

                        $lastRow = Get-LastDataRowFromMap -Cells $cells -ColumnLetters @($nameColLetter, $categoryColLetter, $daysColLetter)

                        if ($lastRow -gt $MaxUserRowsToScan) {
                            Add-ErrorRow $facilityName $monthLabel $file.Name "${nameColLetter}/${categoryColLetter}/${daysColLetter}列付近の最終データ行が $lastRow 行目でした。処理停止防止のため $MaxUserRowsToScan 行目まで読み取ります。"
                            $lastRow = $MaxUserRowsToScan
                        }

                        Write-Host ("  利用者行確認: {0}～{1} 行目（列: {2}/{3}/{4}）" -f $scanStartRow, $lastRow, $nameColLetter, $categoryColLetter, $daysColLetter)

                        for ($row = $scanStartRow; $row -le $lastRow; $row++) {
                            $nameAddress = "${nameColLetter}${row}"
                            $categoryAddress = "${categoryColLetter}${row}"
                            $daysAddress = "${daysColLetter}${row}"

                            if ($cellErrors.ContainsKey($nameAddress) -or $cellErrors.ContainsKey($categoryAddress) -or $cellErrors.ContainsKey($daysAddress)) {
                                Add-ErrorRow $facilityName $monthLabel $file.Name ("{0}行目に数式エラー（{1}）があるため、この行は集計から除外しました。" -f $row, (@($cellErrors[$nameAddress], $cellErrors[$categoryAddress], $cellErrors[$daysAddress]) | Where-Object { $_ }) -join "/")
                                continue
                            }

                            $userName = Get-CellValueFromMap -Cells $cells -Address $nameAddress
                            $category = Get-CellValueFromMap -Cells $cells -Address $categoryAddress
                            $days = Get-CellValueFromMap -Cells $cells -Address $daysAddress

                            if ([string]::IsNullOrWhiteSpace($userName) -or [string]::IsNullOrWhiteSpace($days)) { continue }
                            if ($userName -match "利用者|氏名|名前" -or $days -match "延べ日数|延日数") { continue }
                            # 「合計」「小計」などの集計行を利用者として拾わない
                            if (Test-IsAggregateRowLabel $userName) { continue }

                            if ((Convert-ToNumberOrZero $days) -gt $daysInMonth) {
                                Add-ReviewRow $facilityName $monthLabel $file.Name ("延べ日数が対象月の日数（{0}日）を超えています（利用者: {1}, 値: {2}）" -f $daysInMonth, $userName, $days)
                            }

                            $UserRows += [PSCustomObject]@{
                                "施設名" = $facilityName
                                "棟" = $identity.Unit
                                "提供月" = $monthLabel
                                "種別" = $fileType
                                "ファイル名" = $file.Name
                                "利用者名" = $userName
                                "区分" = Normalize-CategoryText $category
                                "延べ日数" = $days
                            }
                        }

                        # ---- 区分別（K5 / L4:R4 / L5:R5） ----
                        if ($cellErrors.ContainsKey("K5")) {
                            Add-ErrorRow $facilityName $monthLabel $file.Name ("K5（利用回数合計）が数式エラー（{0}）です。区分別集計をスキップしました。" -f $cellErrors["K5"])
                        } else {
                            $totalCount = Get-CellValueFromMap -Cells $cells -Address "K5"
                            $categoryUsageSum = 0.0
                            $categoryFoundAny = $false
                            foreach ($colLetter in @("L", "M", "N", "O", "P", "Q", "R")) {
                                $categoryNameAddress = "${colLetter}4"
                                $usageCountAddress = "${colLetter}5"

                                if ($cellErrors.ContainsKey($categoryNameAddress) -or $cellErrors.ContainsKey($usageCountAddress)) {
                                    Add-ErrorRow $facilityName $monthLabel $file.Name ("区分別集計の{0}列に数式エラーがあるため、この区分は集計から除外しました。" -f $colLetter)
                                    continue
                                }

                                $categoryName = Normalize-CategoryText (Get-CellValueFromMap -Cells $cells -Address $categoryNameAddress)
                                $usageCount = Get-CellValueFromMap -Cells $cells -Address $usageCountAddress

                                if ([string]::IsNullOrWhiteSpace($categoryName) -and [string]::IsNullOrWhiteSpace($usageCount)) { continue }

                                $categoryFoundAny = $true
                                $categoryUsageSum += Convert-ToNumberOrZero $usageCount

                                $CategoryRows += [PSCustomObject]@{
                                    "施設名" = $facilityName
                                    "棟" = $identity.Unit
                                    "フォルダ施設名" = $folderFacilityName
                                    "提供月" = $monthLabel
                                    "種別" = $fileType
                                    "ファイル名" = $file.Name
                                    "区分" = $categoryName
                                    "利用回数" = $usageCount
                                    "利用回数合計" = $totalCount
                                }
                            }

                            if (-not $categoryFoundAny -and -not [string]::IsNullOrWhiteSpace($totalCount)) {
                                Add-ReviewRow $facilityName $monthLabel $file.Name "K5に利用回数合計（$totalCount）がありますが、L4:R4の区分別内訳が見つかりませんでした。レイアウトが異なる可能性があります。"
                            } elseif ($categoryFoundAny -and -not [string]::IsNullOrWhiteSpace($totalCount)) {
                                $totalCountNumber = Convert-ToNumberOrZero $totalCount
                                if ([Math]::Abs($categoryUsageSum - $totalCountNumber) -gt 0.0000001) {
                                    Add-ReviewRow $facilityName $monthLabel $file.Name ("区分別内訳の合計（{0}）とK5の利用回数合計（{1}）が一致しません。レイアウトずれまたは入力誤りの可能性があります。" -f (Format-NumberForCsv $categoryUsageSum), $totalCount)
                                }
                            }
                        }
                    }

                    # ---- 利用者個票（K5がエラーでもここは必ず実行します） ----
                    $detailCount = 0
                    foreach ($workbookSheet in $workbookSheets) {
                        $detailRow = Get-UserRecordDetailRow -FacilityName $facilityName -MonthLabel $monthLabel -FileType $fileType -FileName $file.Name -SheetName $workbookSheet.SheetName -Cells $workbookSheet.Cells
                        if ($null -ne $detailRow) { $UserDetailRows += $detailRow; $detailCount++ }
                    }
                    if ($detailCount -eq 0) {
                        Add-ReviewRow $facilityName $monthLabel $file.Name ("利用者別実績詳細が0件でした。ブック内シート数: {0}。個票シートの判定条件（1〜3行目の実績記録票タイトル、2行目の氏名、日別明細ラベル行）を確認してください。" -f $workbookSheets.Count)
                    }
                    Write-Host ("  利用者別実績詳細: {0} 名分を抽出" -f $detailCount)

                    continue
                }

                # ---- 旧形式（.xls）は Excel COM 経由 ----
                if ($null -eq $excel) {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $excel.DisplayAlerts = $false
                    $excel.EnableEvents = $false
                    $excel.AskToUpdateLinks = $false
                    $excel.AutomationSecurity = 3
                }

                $workbook = $excel.Workbooks.Open($file.FullName, 0, $true)
                $sheet = Get-WorksheetByName -Workbook $workbook -SheetNames $TargetSheetNames

                if ($null -eq $sheet) {
                    Add-ErrorRow $facilityName $monthLabel $file.Name "対象シート「$($TargetSheetNames -join ' / ')」が見つかりません。"
                    continue
                }

                $userColumns = Find-UserTableColumnsFromWorksheet -Worksheet $sheet
                $scanStartRow = 1
                if ($null -eq $userColumns) {
                    $nameColIndex = 2; $categoryColIndex = 3; $daysColIndex = 4
                    Add-ReviewRow $facilityName $monthLabel $file.Name "利用者一覧のヘッダー（利用者名/区分/延べ日数）を自動検出できなかったため、既定のB/C/D列を使用しました。レイアウトを確認してください。"
                } else {
                    $nameColIndex = $userColumns.NameCol
                    $categoryColIndex = $userColumns.CategoryCol
                    $daysColIndex = $userColumns.DaysCol
                    $scanStartRow = $userColumns.HeaderRow + 1
                }

                $nameColLetter = Get-ColumnLetter $nameColIndex
                $categoryColLetter = Get-ColumnLetter $categoryColIndex
                $daysColLetter = Get-ColumnLetter $daysColIndex

                $lastRow = Get-LastDataRowInColumns -Worksheet $sheet -ColumnLetters @($nameColLetter, $categoryColLetter, $daysColLetter)

                if ($lastRow -gt $MaxUserRowsToScan) {
                    Add-ErrorRow $facilityName $monthLabel $file.Name "${nameColLetter}/${categoryColLetter}/${daysColLetter}列付近の最終データ行が $lastRow 行目でした。処理停止防止のため $MaxUserRowsToScan 行目まで読み取ります。"
                    $lastRow = $MaxUserRowsToScan
                }

                Write-Host ("  利用者行確認: {0}～{1} 行目（列: {2}/{3}/{4}）" -f $scanStartRow, $lastRow, $nameColLetter, $categoryColLetter, $daysColLetter)

                for ($row = $scanStartRow; $row -le $lastRow; $row++) {
                    # エラー判定だけ .Text、値は .Value2 を使います。
                    # （.Text は列幅が狭いと "####"、表示形式によっては "6人" 等になり、.xlsx 経路と値がずれます）
                    $nameText = To-Text $sheet.Cells.Item($row, $nameColIndex).Text
                    $categoryText = To-Text $sheet.Cells.Item($row, $categoryColIndex).Text
                    $daysText = To-Text $sheet.Cells.Item($row, $daysColIndex).Text

                    if ((Test-IsExcelErrorText $nameText) -or (Test-IsExcelErrorText $categoryText) -or (Test-IsExcelErrorText $daysText)) {
                        Add-ErrorRow $facilityName $monthLabel $file.Name ("{0}行目に数式エラー（{1}）があるため、この行は集計から除外しました。" -f $row, (@($nameText, $categoryText, $daysText) | Where-Object { Test-IsExcelErrorText $_ }) -join "/")
                        continue
                    }

                    $userName = To-Text $sheet.Cells.Item($row, $nameColIndex).Value2
                    $category = To-Text $sheet.Cells.Item($row, $categoryColIndex).Value2
                    $days = To-Text $sheet.Cells.Item($row, $daysColIndex).Value2

                    if ([string]::IsNullOrWhiteSpace($userName) -or [string]::IsNullOrWhiteSpace($days)) { continue }
                    if ($userName -match "利用者|氏名|名前" -or $days -match "延べ日数|延日数") { continue }
                    if (Test-IsAggregateRowLabel $userName) { continue }

                    if ((Convert-ToNumberOrZero $days) -gt $daysInMonth) {
                        Add-ReviewRow $facilityName $monthLabel $file.Name ("延べ日数が対象月の日数（{0}日）を超えています（利用者: {1}, 値: {2}）" -f $daysInMonth, $userName, $days)
                    }

                    $UserRows += [PSCustomObject]@{
                        "施設名" = $facilityName
                        "棟" = $identity.Unit
                        "提供月" = $monthLabel
                        "種別" = $fileType
                        "ファイル名" = $file.Name
                        "利用者名" = $userName
                        "区分" = Normalize-CategoryText $category
                        "延べ日数" = $days
                    }
                }

                if (Test-IsExcelErrorText (To-Text $sheet.Range("K5").Text)) {
                    Add-ErrorRow $facilityName $monthLabel $file.Name ("K5（利用回数合計）が数式エラー（{0}）です。区分別集計をスキップしました。" -f (To-Text $sheet.Range("K5").Text))
                } else {
                    $totalCount = To-Text $sheet.Range("K5").Value2
                    $categoryUsageSum = 0.0
                    $categoryFoundAny = $false
                    foreach ($col in 12..18) {
                        $categoryNameText = To-Text $sheet.Cells.Item(4, $col).Text
                        $usageCountText = To-Text $sheet.Cells.Item(5, $col).Text

                        if ((Test-IsExcelErrorText $categoryNameText) -or (Test-IsExcelErrorText $usageCountText)) {
                            Add-ErrorRow $facilityName $monthLabel $file.Name ("区分別集計の{0}列目に数式エラーがあるため、この区分は集計から除外しました。" -f $col)
                            continue
                        }

                        $categoryName = Normalize-CategoryText $sheet.Cells.Item(4, $col).Value2
                        $usageCount = To-Text $sheet.Cells.Item(5, $col).Value2

                        if ([string]::IsNullOrWhiteSpace($categoryName) -and [string]::IsNullOrWhiteSpace($usageCount)) { continue }

                        $categoryFoundAny = $true
                        $categoryUsageSum += Convert-ToNumberOrZero $usageCount

                        $CategoryRows += [PSCustomObject]@{
                            "施設名" = $facilityName
                            "棟" = $identity.Unit
                            "フォルダ施設名" = $folderFacilityName
                            "提供月" = $monthLabel
                            "種別" = $fileType
                            "ファイル名" = $file.Name
                            "区分" = $categoryName
                            "利用回数" = $usageCount
                            "利用回数合計" = $totalCount
                        }
                    }

                    if (-not $categoryFoundAny -and -not [string]::IsNullOrWhiteSpace($totalCount)) {
                        Add-ReviewRow $facilityName $monthLabel $file.Name "K5に利用回数合計（$totalCount）がありますが、L4:R4の区分別内訳が見つかりませんでした。レイアウトが異なる可能性があります。"
                    } elseif ($categoryFoundAny -and -not [string]::IsNullOrWhiteSpace($totalCount)) {
                        $totalCountNumber = Convert-ToNumberOrZero $totalCount
                        if ([Math]::Abs($categoryUsageSum - $totalCountNumber) -gt 0.0000001) {
                            Add-ReviewRow $facilityName $monthLabel $file.Name ("区分別内訳の合計（{0}）とK5の利用回数合計（{1}）が一致しません。レイアウトずれまたは入力誤りの可能性があります。" -f (Format-NumberForCsv $categoryUsageSum), $totalCount)
                        }
                    }
                }

                Add-ReviewRow $facilityName $monthLabel $file.Name "旧形式（.xls）のファイルのため、利用者別実績詳細（加算・入院/外泊等）は抽出していません。新形式（.xlsx/.xlsm）に変換すると抽出できます。"
            }
            catch {
                Add-ErrorRow $facilityName $monthLabel $file.Name $_.Exception.Message
            }
            finally {
                if ($null -ne $workbook) { $workbook.Close($false) }
                Release-ComObject $sheet
                Release-ComObject $workbook
            }
        }

        }
        catch {
            Add-ErrorRow $folderFacilityName $monthLabel "" ("施設フォルダの処理中に予期しないエラーが発生したため、この施設・月をスキップしました。パス: {0} / エラー: {1}" -f $monthFolder.Path, $_.Exception.Message)
            continue
        }
    }
}
finally {
    if ($null -ne $excel) { $excel.Quit() }
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ==============================
# 集計・出力
# ==============================
$outputPath = Join-Path $OutputDir "実績記録表_抽出結果.xlsx"

$UserRows = @($UserRows)
$UserDetailRows = @($UserDetailRows)
$CategoryRows = @($CategoryRows)
$ErrorRows = @($ErrorRows)
$ReviewRows = @($ReviewRows)
$SkippedRows = @($SkippedRows)

$CategoryRowsWithTotals = @(Add-AdmissionTotalRows -Rows $CategoryRows)
$CategorySummaryRows = @(Get-CategoryRowsForSummaryWorkbook -Rows $CategoryRows)

$FacilityGrandTotals = @(
    $CategorySummaryRows |
        Group-Object { "{0}`t{1}" -f $_."施設名", $_."提供月" } |
        ForEach-Object {
            $firstRow = $_.Group | Select-Object -First 1
            $total = 0.0
            foreach ($row in $_.Group) { $total += Convert-ToNumberOrZero $row."合計" }
            [PSCustomObject]@{
                "施設名" = $firstRow."施設名"
                "提供月" = $firstRow."提供月"
                "合計利用回数" = Format-NumberForCsv $total
            }
        } |
        Sort-Object "施設名", "提供月"
)

$AdmissionMatrixRows = @(Get-FacilityCategoryMatrixRows -Rows $CategoryRows -IncludeTypes $AdmissionTypes)
$ShortStayMatrixRows = @(Get-FacilityCategoryMatrixRows -Rows $CategoryRows -IncludeTypes $ShortStayTypes)
$CombinedMatrixRows = @(Get-FacilityCategoryMatrixRows -Rows $CategoryRows -IncludeTypes ($AdmissionTypes + $ShortStayTypes))

# ---- 検算 ----
# 元データ（区分別行）の利用回数合計と、各シートの合計が一致するかを必ず確認します。
$sourceTotalAll = Get-RowsTotal -Rows $CategoryRows -ColumnName "利用回数"
$sourceTotalAdmission = Get-RowsTotal -Rows @($CategoryRows | Where-Object { $AdmissionTypes -contains $_."種別" }) -ColumnName "利用回数"
$sourceTotalShortStay = Get-RowsTotal -Rows @($CategoryRows | Where-Object { $ShortStayTypes -contains $_."種別" }) -ColumnName "利用回数"

function New-ReconciliationRow {
    param([string]$Name, [double]$SourceTotal, [double]$SheetTotal)

    $difference = $SheetTotal - $SourceTotal
    return [PSCustomObject]@{
        "検算対象" = $Name
        "元データ合計" = Format-NumberForCsv $SourceTotal
        "シート合計" = Format-NumberForCsv $SheetTotal
        "差分" = Format-NumberForCsv $difference
        "判定" = if ([Math]::Abs($difference) -lt 0.0000001) { "OK" } else { "要確認" }
    }
}

$ReconciliationRows = @(
    New-ReconciliationRow "施設別・区分別合計（入居）" $sourceTotalAdmission (Get-MatrixTotal -Rows $AdmissionMatrixRows)
    New-ReconciliationRow "施設別・区分別合計（短期）" $sourceTotalShortStay (Get-MatrixTotal -Rows $ShortStayMatrixRows)
    New-ReconciliationRow "施設別・区分別合計（入居＋短期）" $sourceTotalAll (Get-MatrixTotal -Rows $CombinedMatrixRows)
    New-ReconciliationRow "施設別合計" $sourceTotalAll (Get-RowsTotal -Rows $CategorySummaryRows -ColumnName "合計")
    New-ReconciliationRow "サマリー（施設別合計利用回数）" $sourceTotalAll (Get-RowsTotal -Rows $FacilityGrandTotals -ColumnName "合計利用回数")
)

$reconciliationNg = @($ReconciliationRows | Where-Object { $_."判定" -ne "OK" })
foreach ($ngRow in $reconciliationNg) {
    Add-ReviewRow "" ($targetMonthLabels -join ", ") "" ("検算NG: シート「{0}」の合計（{1}）が元データの合計（{2}）と一致しません（差分 {3}）。" -f $ngRow."検算対象", $ngRow."シート合計", $ngRow."元データ合計", $ngRow."差分")
}
$ReviewRows = @($ReviewRows)

$summaryInfo = [PSCustomObject]@{
    "実行日時" = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    "対象月" = ($targetMonthLabels -join ", ")
    "対象フォルダ数" = $monthFolders.Count
    "集計した施設数" = @($CategoryRows | ForEach-Object { $_."施設名" } | Sort-Object -Unique).Count
    "利用者別データ件数" = $UserRows.Count
    "利用者別実績詳細件数" = $UserDetailRows.Count
    "区分別データ件数（入居合計行を含む）" = $CategoryRowsWithTotals.Count
    "利用回数の総合計（元データ）" = Format-NumberForCsv $sourceTotalAll
    "検算結果" = if ($reconciliationNg.Count -eq 0) { "OK（全シート一致）" } else { ("要確認（{0}件不一致）" -f $reconciliationNg.Count) }
    "エラー件数" = $ErrorRows.Count
    "確認対象件数" = $ReviewRows.Count
    "対象外ファイル件数" = $SkippedRows.Count
}

$savedOutputPath = Write-ResultWorkbook `
    -SummaryInfo $summaryInfo `
    -ReconciliationRows $ReconciliationRows `
    -FacilityGrandTotals $FacilityGrandTotals `
    -AdmissionMatrixRows $AdmissionMatrixRows `
    -ShortStayMatrixRows $ShortStayMatrixRows `
    -CombinedMatrixRows $CombinedMatrixRows `
    -CategorySummaryRows $CategorySummaryRows `
    -CategoryDetailRows $CategoryRowsWithTotals `
    -UserRows $UserRows `
    -UserDetailRows $UserDetailRows `
    -ErrorRows $ErrorRows `
    -ReviewRows $ReviewRows `
    -SkippedRows $SkippedRows `
    -Path $outputPath

Write-Host ""
Write-Host "抽出が完了しました。"
Write-Host ("利用者別: {0} 件" -f $UserRows.Count)
Write-Host ("利用者別実績詳細: {0} 件" -f $UserDetailRows.Count)
Write-Host ("区分別: {0} 件（入居合計行を含む）" -f $CategoryRowsWithTotals.Count)
Write-Host ("施設別合計: {0} 件" -f $CategorySummaryRows.Count)
Write-Host ("エラー: {0} 件" -f $ErrorRows.Count)
Write-Host ("確認対象: {0} 件" -f $ReviewRows.Count)
Write-Host ("対象外ファイル: {0} 件" -f $SkippedRows.Count)
Write-Host ""
if ($reconciliationNg.Count -eq 0) {
    Write-Host "検算: OK（元データと全シートの合計が一致しました）" -ForegroundColor Green
} else {
    Write-Host ("検算: 要確認（{0}件のシートで合計が一致しません。サマリーシートの検算表と確認対象一覧を見てください）" -f $reconciliationNg.Count) -ForegroundColor Yellow
}
Write-Host ""
Write-Host "出力先:"
Write-Host $savedOutputPath
