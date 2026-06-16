Attribute VB_Name = "ShiftChecker"
' ===================================================================
' 週次シフト分析マクロ  ShiftChecker.bas
'
' 【使い方】
'   1. このファイルをExcelのVBAエディタ(Alt+F11)で
'      「ファイル」→「ファイルのインポート」でインポート
'   2. 開発タブ or VBAエディタから RunShiftCheck を実行
'   3. 年・月を入力すると全施設を自動解析
'
' 【保存先ブック(.xlsm)】
'   C:\Users\k-takayama\OneDrive\デスクトップ\📝AM業務アプリ📝\週次シフト分析\
'
' 【シフトデータ参照先 (Dropbox)】
'   C:\Users\k-takayama\Bihonest Corp. Dropbox\17.ラシエル\
'       00.ラシエル_General\1 各施設シフト管理\
' ===================================================================
Option Explicit

' --- 定数 ---
Private Const BASE_PATH       As String  = "C:\Users\k-takayama\Bihonest Corp. Dropbox\17.ラシエル\00.ラシエル_General\1 各施設シフト管理\"
Private Const MAX_RUNS        As Integer = 5
Private Const RESULT_SHEET    As String  = "チェック結果"
Private Const HIST_FAC_SHEET  As String  = "履歴_施設"
Private Const HIST_STAF_SHEET As String  = "履歴_職員"

' AI列=35, CQ列=95
Private Const VAC_START_COL   As Integer = 35
Private Const VAC_END_COL     As Integer = 95

' ===================================================================
'  メインエントリーポイント
' ===================================================================
Public Sub RunShiftCheck()

    ' --- 年月の入力 ---
    Dim year4  As String
    Dim month2 As String

    year4 = InputBox("対象年を4桁で入力してください（例：2026）", "週次シフト分析", Format(Year(Now), "0000"))
    If year4 = "" Then Exit Sub
    If Len(year4) <> 4 Or Not IsNumeric(year4) Then
        MsgBox "年は4桁の数字で入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    month2 = InputBox("対象月を2桁で入力してください（例：06）", "週次シフト分析", Format(Month(Now), "00"))
    If month2 = "" Then Exit Sub
    If Len(month2) <> 2 Or Not IsNumeric(month2) _
       Or CInt(month2) < 1 Or CInt(month2) > 12 Then
        MsgBox "月は01〜12の2桁数字で入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    Dim yymm    As String
    Dim runDt   As Date
    yymm  = Right(year4, 2) & month2
    runDt = Now()

    ' --- シート準備 ---
    EnsureSheet RESULT_SHEET
    EnsureSheet HIST_FAC_SHEET
    EnsureSheet HIST_STAF_SHEET

    ' --- 実行ID採番 ---
    Dim runId As Long
    runId = NextRunId()

    ' --- 施設フォルダ列挙 ---
    Dim folders()    As String
    Dim facCount     As Integer
    facCount = ScanFacilityFolders(BASE_PATH, folders)

    If facCount = 0 Then
        MsgBox "施設フォルダが見つかりませんでした。" & vbCrLf & BASE_PATH, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False
    Application.StatusBar      = "シフト解析中..."

    ' 結果格納配列 (施設数 x 項目数)
    ' 列定義:
    '  0=フォルダ名, 1=施設名, 2=状態("OK"/"FILE_NOT_FOUND"/"ERROR:...")
    '  3=DL3, 4=FE79, 5=FE82, 6=EV79, 7=EV78
    '  8=DY24, 9=EA19, 10=EA22
    '  11=DY40, 12=EA30, 13=EA36
    '  14=DU75, 15=BU73, 16=BU75, 17=realTotal
    Dim res() As Variant
    ReDim res(0 To facCount - 1, 0 To 20)

    Dim i As Integer
    For i = 0 To facCount - 1
        Dim facName  As String
        Dim yearDir  As String
        Dim filePath As String

        facName  = ExtractFacilityName(folders(i))
        yearDir  = BASE_PATH & folders(i) & "\" & year4 & "年\"
        filePath = FindMonthFile(yearDir, yymm)

        res(i, 0) = folders(i)
        res(i, 1) = facName

        Application.StatusBar = "解析中: " & facName

        If filePath = "" Then
            res(i, 2) = "FILE_NOT_FOUND"
        Else
            ProcessShiftFile filePath, runId, runDt, facName, res, i
        End If
    Next i

    ' --- 古い履歴を削除 ---
    PruneHistory runId

    ' --- 結果シート出力 ---
    WriteResultSheet year4, month2, res, facCount, runId

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
    Application.StatusBar      = False

    ThisWorkbook.Sheets(RESULT_SHEET).Activate
    MsgBox year4 & "年" & month2 & "月 シフトチェック完了" & vbCrLf & _
           facCount & "施設を処理しました。", vbInformation, "週次シフト分析"
End Sub

' ===================================================================
'  施設フォルダ列挙
' ===================================================================
Private Function ScanFacilityFolders(basePath As String, ByRef folders() As String) As Integer
    Dim fso    As Object
    Dim folder As Object
    Dim sub_   As Object
    Dim count  As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(basePath) Then
        ScanFacilityFolders = 0
        Exit Function
    End If

    Set folder = fso.GetFolder(basePath)
    ReDim folders(0 To 100)
    count = 0

    For Each sub_ In folder.SubFolders
        folders(count) = sub_.Name
        count = count + 1
    Next sub_

    ScanFacilityFolders = count
End Function

Private Function ExtractFacilityName(folderName As String) As String
    ' "16.南中丸" → "南中丸"
    Dim dotPos As Integer
    dotPos = InStr(folderName, ".")
    If dotPos >= 1 And dotPos <= 3 Then
        ExtractFacilityName = Mid(folderName, dotPos + 1)
    Else
        ExtractFacilityName = folderName
    End If
End Function

Private Function FindMonthFile(folderPath As String, yymm As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        FindMonthFile = ""
        Exit Function
    End If

    ' YYMM で始まる .xlsx を探す
    Dim fn As String
    fn = Dir(folderPath & yymm & "*.xlsx")
    If fn <> "" Then
        FindMonthFile = folderPath & fn
    Else
        FindMonthFile = ""
    End If
End Function

' ===================================================================
'  1ファイルの処理（セル読み取り + 履歴保存）
' ===================================================================
Private Sub ProcessShiftFile(filePath As String, runId As Long, runDt As Date, _
                              facName As String, ByRef res() As Variant, idx As Integer)
    Dim wb As Workbook
    On Error GoTo ErrHandler

    Set wb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)

    ' 管理用シートを検索
    Dim ws      As Worksheet
    Dim sh      As Object
    Set ws = Nothing
    For Each sh In wb.Sheets
        If InStr(sh.Name, "管理用シート") > 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh
    If ws Is Nothing Then Set ws = wb.Sheets(1)

    ' ---- セル値取得 ----
    Dim dl3  As Variant : dl3  = CV(ws, "DL3")
    Dim fe79 As Variant : fe79 = CV(ws, "FE79")
    Dim fe82 As Variant : fe82 = CV(ws, "FE82")
    Dim ev79 As Variant : ev79 = CV(ws, "EV79")
    Dim ev78 As Variant : ev78 = CV(ws, "EV78")
    Dim ea19 As Variant : ea19 = CV(ws, "EA19")
    Dim ea22 As Variant : ea22 = CV(ws, "EA22")
    Dim dy24 As Double  : dy24 = ToN(CV(ws, "DY24"))
    Dim ea30 As Variant : ea30 = CV(ws, "EA30")
    Dim ea36 As Variant : ea36 = CV(ws, "EA36")
    Dim dy40 As Double  : dy40 = ToN(CV(ws, "DY40"))
    Dim du75 As Variant : du75 = CV(ws, "DU75")
    Dim bu73 As Variant : bu73 = CV(ws, "BU73")
    Dim bu75 As Variant : bu75 = CV(ws, "BU75")
    Dim realTotal As Double : realTotal = dy24 + dy40

    ' 結果配列に格納
    res(idx, 2)  = "OK"
    res(idx, 3)  = dl3 : res(idx, 4)  = fe79 : res(idx, 5)  = fe82
    res(idx, 6)  = ev79 : res(idx, 7)  = ev78
    res(idx, 8)  = dy24 : res(idx, 9)  = ea19 : res(idx, 10) = ea22
    res(idx, 11) = dy40 : res(idx, 12) = ea30 : res(idx, 13) = ea36
    res(idx, 14) = du75 : res(idx, 15) = bu73 : res(idx, 16) = bu75
    res(idx, 17) = realTotal

    ' 施設履歴保存
    SaveFacHistory runId, runDt, facName, res, idx

    ' ---- 職員行ループ (行9〜58) ----
    Dim r As Integer
    For r = 9 To 58
        Dim bv As Variant : bv  = CV(ws, "B"  & r)
        Dim mv As Variant : mv  = CV(ws, "M"  & r)
        Dim cuv As Variant : cuv = CV(ws, "CU" & r)
        Dim duv As Variant : duv = CV(ws, "DU" & r)
        Dim dlv As Variant : dlv = CV(ws, "DL" & r)
        Dim aev As Variant : aev = CV(ws, "AE" & r)

        If bv = "" And mv = "" And cuv = "" And duv = "" And dlv = "" Then GoTo NextRow

        ' AI〜CQ 列の公休・有休カウント
        Dim vacCnt  As Integer : vacCnt  = 0
        Dim paidCnt As Integer : paidCnt = 0
        Dim c As Integer
        For c = VAC_START_COL To VAC_END_COL
            Dim cellVal As Variant
            cellVal = ws.Cells(r, c).Value
            If Not IsEmpty(cellVal) And cellVal <> "" Then
                Dim cs As String
                cs = Trim(CStr(cellVal))
                Select Case cs
                    Case "休", "休み", "公", "公休", "希休", "希", "希望休"
                        vacCnt = vacCnt + 1
                    Case Else
                        If InStr(cs, "有") > 0 Then paidCnt = paidCnt + 1
                End Select
            End If
        Next c

        SaveStaffHistory runId, runDt, facName, r, bv, mv, cuv, duv, dlv, aev, dl3, vacCnt, paidCnt

NextRow:
    Next r

    wb.Close SaveChanges:=False
    Exit Sub

ErrHandler:
    res(idx, 2) = "ERROR: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' セル値取得ヘルパー
Private Function CV(ws As Worksheet, addr As String) As Variant
    Dim v As Variant
    v = ws.Range(addr).Value
    If IsEmpty(v) Then CV = "" Else CV = v
End Function

' 数値変換ヘルパー
Private Function ToN(v As Variant) As Double
    If IsNumeric(v) Then ToN = CDbl(v) Else ToN = 0
End Function

' ===================================================================
'  履歴シートへの保存
' ===================================================================
Private Sub SaveFacHistory(runId As Long, runDt As Date, facName As String, _
                            ByRef res() As Variant, idx As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HIST_FAC_SHEET)

    If ws.Cells(1, 1).Value = "" Then
        Dim hdr1 As Variant
        hdr1 = Array("実行日時", "実行ID", "施設名", "所定時間(DL3)", _
                     "支援員基準(FE79)", "世話人基準(FE82)", "追加必要人数(EV79)", "体制加算判定(EV78)", _
                     "支援員実配置(DY24)", "支援員日中換算(EA19)", "支援員超過換算(EA22)", _
                     "世話人実配置(DY40)", "世話人日中換算(EA30)", "世話人超過換算(EA36)", _
                     "7.5:1基準人員(DU75)", "現在基準値(BU73)", "過不足(BU75)", "実配置合計")
        Dim hi As Integer
        For hi = 0 To UBound(hdr1)
            ws.Cells(1, hi + 1).Value = hdr1(hi)
        Next hi
        ws.Rows(1).Font.Bold = True
    End If

    Dim nr As Long
    nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nr, 1).Value  = runDt
    ws.Cells(nr, 2).Value  = runId
    ws.Cells(nr, 3).Value  = facName
    ws.Cells(nr, 4).Value  = res(idx, 3)   ' dl3
    ws.Cells(nr, 5).Value  = res(idx, 4)   ' fe79
    ws.Cells(nr, 6).Value  = res(idx, 5)   ' fe82
    ws.Cells(nr, 7).Value  = res(idx, 6)   ' ev79
    ws.Cells(nr, 8).Value  = res(idx, 7)   ' ev78
    ws.Cells(nr, 9).Value  = res(idx, 8)   ' dy24
    ws.Cells(nr, 10).Value = res(idx, 9)   ' ea19
    ws.Cells(nr, 11).Value = res(idx, 10)  ' ea22
    ws.Cells(nr, 12).Value = res(idx, 11)  ' dy40
    ws.Cells(nr, 13).Value = res(idx, 12)  ' ea30
    ws.Cells(nr, 14).Value = res(idx, 13)  ' ea36
    ws.Cells(nr, 15).Value = res(idx, 14)  ' du75
    ws.Cells(nr, 16).Value = res(idx, 15)  ' bu73
    ws.Cells(nr, 17).Value = res(idx, 16)  ' bu75
    ws.Cells(nr, 18).Value = res(idx, 17)  ' realTotal
End Sub

Private Sub SaveStaffHistory(runId As Long, runDt As Date, facName As String, rowNum As Integer, _
                              bv As Variant, mv As Variant, cuv As Variant, duv As Variant, _
                              dlv As Variant, aev As Variant, dl3 As Variant, _
                              vacCnt As Integer, paidCnt As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HIST_STAF_SHEET)

    If ws.Cells(1, 1).Value = "" Then
        Dim hdr2 As Variant
        hdr2 = Array("実行日時", "実行ID", "施設名", "行番号", "職種", "氏名", _
                     "総労働時間(CU)", "日中時間(DU)", "日中配置(DL)", "常勤区分(AE)", _
                     "所定時間(DL3)", "公休数", "有休数")
        Dim hi As Integer
        For hi = 0 To UBound(hdr2)
            ws.Cells(1, hi + 1).Value = hdr2(hi)
        Next hi
        ws.Rows(1).Font.Bold = True
    End If

    Dim nr As Long
    nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nr, 1).Value  = runDt
    ws.Cells(nr, 2).Value  = runId
    ws.Cells(nr, 3).Value  = facName
    ws.Cells(nr, 4).Value  = rowNum
    ws.Cells(nr, 5).Value  = bv
    ws.Cells(nr, 6).Value  = mv
    ws.Cells(nr, 7).Value  = cuv
    ws.Cells(nr, 8).Value  = duv
    ws.Cells(nr, 9).Value  = dlv
    ws.Cells(nr, 10).Value = aev
    ws.Cells(nr, 11).Value = dl3
    ws.Cells(nr, 12).Value = vacCnt
    ws.Cells(nr, 13).Value = paidCnt
End Sub

' ===================================================================
'  古い履歴削除（MAX_RUNS を超えた実行分を除去）
' ===================================================================
Private Sub PruneHistory(currentRunId As Long)
    If currentRunId <= MAX_RUNS Then Exit Sub
    Dim cutoff As Long
    cutoff = currentRunId - MAX_RUNS
    PruneSheet HIST_FAC_SHEET,  cutoff
    PruneSheet HIST_STAF_SHEET, cutoff
End Sub

Private Sub PruneSheet(sheetName As String, cutoffId As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = lastRow To 2 Step -1
        If IsNumeric(ws.Cells(r, 2).Value) Then
            If CLng(ws.Cells(r, 2).Value) <= cutoffId Then
                ws.Rows(r).Delete
            End If
        End If
    Next r
End Sub

' ===================================================================
'  結果シート出力
' ===================================================================
Private Sub WriteResultSheet(year4 As String, month2 As String, _
                              ByRef res() As Variant, facCount As Integer, runId As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(RESULT_SHEET)
    ws.Cells.Clear
    ws.Tab.Color = RGB(46, 204, 113)

    ' --- タイトル ---
    With ws.Cells(1, 1)
        .Value = year4 & "年" & month2 & "月　週次シフト分析チェック結果"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    ws.Cells(2, 1).Value = "実行日時: " & Format(Now(), "yyyy/mm/dd hh:mm") & _
                           "　　実行ID: " & runId

    ' ==============================================
    '  ❶ 施設サマリー表
    ' ==============================================
    Dim sRow As Long
    sRow = 4

    ws.Cells(sRow, 1).Value = "【❶ 施設サマリー】"
    ws.Cells(sRow, 1).Font.Bold = True
    ws.Cells(sRow, 1).Font.Size = 12
    sRow = sRow + 1

    ' ヘッダー
    Dim sh1 As Variant
    sh1 = Array("施設名", "所定時間", "体制加算判定", "実配置合計", "過不足(BU75)", _
                "支援員実配置", "世話人実配置", "前回比:実配置", "前回比:過不足", "ステータス")
    Dim h As Integer
    For h = 0 To UBound(sh1)
        With ws.Cells(sRow, h + 1)
            .Value = sh1(h)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(44, 62, 80)
        End With
    Next h
    sRow = sRow + 1

    Dim i As Integer
    For i = 0 To facCount - 1
        ws.Cells(sRow, 1).Value = res(i, 1)  ' 施設名

        If res(i, 2) = "OK" Then
            Dim ev78Str As String : ev78Str = CStr(res(i, 7))
            Dim bu75Val As Double : bu75Val = ToN(res(i, 16))
            Dim rtVal   As Double : rtVal   = ToN(res(i, 17))

            ws.Cells(sRow, 2).Value = res(i, 3)      ' dl3
            ws.Cells(sRow, 3).Value = ev78Str         ' ev78
            ws.Cells(sRow, 4).Value = rtVal           ' realTotal
            ws.Cells(sRow, 5).Value = bu75Val         ' bu75
            ws.Cells(sRow, 6).Value = res(i, 8)       ' dy24
            ws.Cells(sRow, 7).Value = res(i, 11)      ' dy40

            ' EV78 警告
            If InStr(ev78Str, "未達") > 0 Then
                ws.Cells(sRow, 3).Font.Color = RGB(220, 38, 38)
                ws.Cells(sRow, 3).Font.Bold  = True
                ws.Cells(sRow, 3).Interior.Color = RGB(254, 226, 226)
            End If

            ' BU75 警告
            If bu75Val >= 0.1 Then
                ws.Cells(sRow, 5).Font.Color = RGB(220, 38, 38)
                ws.Cells(sRow, 5).Font.Bold  = True
                ws.Cells(sRow, 5).Interior.Color = RGB(254, 226, 226)
            End If

            ' --- 前回比較 ---
            Dim prevFac As Variant
            prevFac = GetPrevFacData(res(i, 1), runId)

            If Not IsEmpty(prevFac(0)) Then
                ' 実配置合計の差分
                Dim rtDiff As Double
                rtDiff = rtVal - ToN(prevFac(0))
                ws.Cells(sRow, 8).Value = FormatDiff(rtDiff, "0.0")
                ws.Cells(sRow, 8).Font.Color = DiffColor(rtDiff)

                ' 過不足の差分
                Dim buDiff As Double
                buDiff = bu75Val - ToN(prevFac(1))
                ws.Cells(sRow, 9).Value = FormatDiff(buDiff, "0.00")
                ws.Cells(sRow, 9).Font.Color = DiffColor(buDiff)
            Else
                ws.Cells(sRow, 8).Value = "(初回)"
                ws.Cells(sRow, 8).Font.Color = RGB(148, 163, 184)
                ws.Cells(sRow, 9).Value = "(初回)"
                ws.Cells(sRow, 9).Font.Color = RGB(148, 163, 184)
            End If

            ' ステータス
            Dim status As String : status = ""
            If InStr(ev78Str, "未達") > 0 Then status = status & "⚠ 体制加算未達  "
            If bu75Val >= 0.1 Then              status = status & "⚠ 過不足超過  "
            If status = "" Then
                status = "OK"
                ws.Cells(sRow, 10).Font.Color = RGB(39, 174, 96)
            Else
                ws.Cells(sRow, 10).Font.Color = RGB(220, 38, 38)
                ws.Cells(sRow, 10).Font.Bold  = True
            End If
            ws.Cells(sRow, 10).Value = status

        ElseIf res(i, 2) = "FILE_NOT_FOUND" Then
            ws.Cells(sRow, 2).Value = "ファイル未発見"
            ws.Cells(sRow, 2).Font.Color = RGB(148, 163, 184)
            ws.Cells(sRow, 10).Value = "ファイル未発見"
            ws.Cells(sRow, 10).Font.Color = RGB(180, 100, 0)
        Else
            ws.Cells(sRow, 2).Value = res(i, 2)
            ws.Cells(sRow, 10).Value = "ERROR"
            ws.Cells(sRow, 10).Font.Color = RGB(220, 38, 38)
        End If

        ' ゼブラ縞
        If (i Mod 2) = 1 Then
            ws.Range(ws.Cells(sRow, 1), ws.Cells(sRow, 10)).Interior.Color = RGB(248, 250, 252)
        End If

        sRow = sRow + 1
    Next i

    ' 罫線
    Dim summaryRange As Range
    Set summaryRange = ws.Range(ws.Cells(5, 1), ws.Cells(sRow - 1, 10))
    summaryRange.Borders.LineStyle = xlContinuous
    summaryRange.Borders.Color = RGB(203, 213, 225)

    ' ==============================================
    '  ❷ 職員別変更点リスト
    ' ==============================================
    sRow = sRow + 2

    ws.Cells(sRow, 1).Value = "【❷ 職員別変更点】　※前回実行との差分のみ表示"
    ws.Cells(sRow, 1).Font.Bold = True
    ws.Cells(sRow, 1).Font.Size = 12
    sRow = sRow + 1

    Dim sh2 As Variant
    sh2 = Array("施設名", "氏名", "職種", "変更項目", "前回値", "今回値", "増減")
    For h = 0 To UBound(sh2)
        With ws.Cells(sRow, h + 1)
            .Value = sh2(h)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(52, 73, 94)
        End With
    Next h
    sRow = sRow + 1

    Dim changeCount As Integer : changeCount = 0

    For i = 0 To facCount - 1
        If res(i, 2) <> "OK" Then GoTo NextFac

        Dim changes As Collection
        Set changes = GetStaffChanges(res(i, 1), runId)

        Dim ch As Variant
        For Each ch In changes
            ws.Cells(sRow, 1).Value = ch(0)  ' 施設名
            ws.Cells(sRow, 2).Value = ch(1)  ' 氏名
            ws.Cells(sRow, 3).Value = ch(2)  ' 職種
            ws.Cells(sRow, 4).Value = ch(3)  ' 変更項目
            ws.Cells(sRow, 5).Value = ch(4)  ' 前回値
            ws.Cells(sRow, 6).Value = ch(5)  ' 今回値
            ws.Cells(sRow, 7).Value = ch(6)  ' 増減

            Select Case CStr(ch(3))
                Case "新規追加"
                    ws.Cells(sRow, 4).Font.Color = RGB(39, 174, 96)
                    ws.Cells(sRow, 4).Font.Bold  = True
                    ws.Rows(sRow).Interior.Color = RGB(240, 253, 244)
                Case "いなくなった"
                    ws.Cells(sRow, 4).Font.Color = RGB(231, 76, 60)
                    ws.Cells(sRow, 4).Font.Bold  = True
                    ws.Rows(sRow).Interior.Color = RGB(254, 242, 242)
                Case "総労働時間"
                    ws.Cells(sRow, 7).Font.Color = DiffColor(ToN(ch(6)))
            End Select

            sRow = sRow + 1
            changeCount = changeCount + 1
        Next ch

NextFac:
    Next i

    If changeCount = 0 Then
        With ws.Cells(sRow, 1)
            .Value = "前回から変更はありませんでした。"
            .Font.Color = RGB(100, 116, 139)
            .Font.Italic = True
        End With
    End If

    ws.Columns.AutoFit
    ws.Rows.AutoFit

    ' ウィンドウ枠の固定（サマリー先頭行）
    ws.Activate
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = False
End Sub

' ===================================================================
'  前回施設データ取得
'  戻り値: (0)=realTotal, (1)=bu75  が Empty なら前回データなし
' ===================================================================
Private Function GetPrevFacData(facName As String, currentRunId As Long) As Variant
    Dim result(1) As Variant
    result(0) = Empty : result(1) = Empty

    Dim prevId As Long : prevId = currentRunId - 1
    If prevId < 1 Then
        GetPrevFacData = result
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HIST_FAC_SHEET)
    Dim lastRow As Long : lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, 2).Value = prevId And CStr(ws.Cells(r, 3).Value) = facName Then
            result(0) = ws.Cells(r, 18).Value  ' realTotal
            result(1) = ws.Cells(r, 17).Value  ' bu75
            Exit For
        End If
    Next r

    GetPrevFacData = result
End Function

' ===================================================================
'  職員別変更点取得（施設単位）
' ===================================================================
Private Function GetStaffChanges(facName As String, currentRunId As Long) As Collection
    Dim coll As New Collection

    Dim prevId As Long : prevId = currentRunId - 1
    If prevId < 1 Then
        Set GetStaffChanges = coll
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HIST_STAF_SHEET)
    Dim lastRow As Long : lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim prevDict As Object : Set prevDict = CreateObject("Scripting.Dictionary")
    Dim currDict As Object : Set currDict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 3).Value) <> facName Then GoTo NextSR

        Dim rid As Long : rid = CLng(ws.Cells(r, 2).Value)
        Dim nm  As String : nm = Trim(CStr(ws.Cells(r, 6).Value))
        If nm = "" Or nm = "-" Then GoTo NextSR

        Dim va(7) As Variant
        va(0) = ws.Cells(r, 5).Value   ' 職種
        va(1) = ws.Cells(r, 7).Value   ' CU 総労働時間
        va(2) = ws.Cells(r, 8).Value   ' DU 日中時間
        va(3) = ws.Cells(r, 10).Value  ' AE 常勤区分
        va(4) = ws.Cells(r, 11).Value  ' DL3 所定時間
        va(5) = ws.Cells(r, 12).Value  ' 公休数
        va(6) = ws.Cells(r, 13).Value  ' 有休数
        va(7) = ws.Cells(r, 9).Value   ' DL 日中配置

        Dim vaVar As Variant : vaVar = va   ' 配列→Variant

        If rid = prevId Then
            prevDict(nm) = vaVar
        ElseIf rid = currentRunId Then
            currDict(nm) = vaVar
        End If

NextSR:
    Next r

    Dim k As Variant

    ' 今回にいて前回にいない → 新規追加
    For Each k In currDict.Keys
        If Not prevDict.Exists(k) Then
            coll.Add MakeRow(facName, CStr(k), CStr(currDict(k)(0)), "新規追加", "(なし)", CStr(k), "")
        End If
    Next k

    ' 前回にいて今回にいない → いなくなった
    For Each k In prevDict.Keys
        If Not currDict.Exists(k) Then
            coll.Add MakeRow(facName, CStr(k), CStr(prevDict(k)(0)), "いなくなった", CStr(k), "(なし)", "")
        End If
    Next k

    ' 両方いる → 項目比較
    For Each k In currDict.Keys
        If Not prevDict.Exists(k) Then GoTo NextCK

        Dim cur As Variant : cur = currDict(k)
        Dim prv As Variant : prv = prevDict(k)

        ' 総労働時間
        If IsNumeric(cur(1)) And IsNumeric(prv(1)) Then
            Dim cuDiff As Double : cuDiff = CDbl(cur(1)) - CDbl(prv(1))
            If Abs(cuDiff) >= 0.5 Then
                coll.Add MakeRow(facName, CStr(k), CStr(cur(0)), "総労働時間", _
                                 Format(prv(1), "0.0") & "h", _
                                 Format(cur(1), "0.0") & "h", _
                                 FormatDiff(cuDiff, "0.0") & "h")
            End If
        End If

        ' 公休数
        If IsNumeric(cur(5)) And IsNumeric(prv(5)) Then
            Dim vDiff As Long : vDiff = CLng(cur(5)) - CLng(prv(5))
            If vDiff <> 0 Then
                coll.Add MakeRow(facName, CStr(k), CStr(cur(0)), "公休数", _
                                 prv(5) & "日", cur(5) & "日", FormatDiff(vDiff, "0") & "日")
            End If
        End If

        ' 有休数
        If IsNumeric(cur(6)) And IsNumeric(prv(6)) Then
            Dim pDiff As Long : pDiff = CLng(cur(6)) - CLng(prv(6))
            If pDiff <> 0 Then
                coll.Add MakeRow(facName, CStr(k), CStr(cur(0)), "有休数", _
                                 prv(6) & "日", cur(6) & "日", FormatDiff(pDiff, "0") & "日")
            End If
        End If

        ' 常勤区分の変化
        If CStr(cur(3)) <> CStr(prv(3)) And CStr(prv(3)) <> "" Then
            coll.Add MakeRow(facName, CStr(k), CStr(cur(0)), "常勤区分", _
                             CStr(prv(3)), CStr(cur(3)), "")
        End If

NextCK:
    Next k

    Set GetStaffChanges = coll
End Function

' 変更行配列を作成するヘルパー
Private Function MakeRow(fac As String, nm As String, job As String, item As String, _
                          prev As String, curr As String, diff As String) As Variant
    Dim row(6) As Variant
    row(0) = fac : row(1) = nm : row(2) = job
    row(3) = item : row(4) = prev : row(5) = curr : row(6) = diff
    MakeRow = row
End Function

' ===================================================================
'  ユーティリティ
' ===================================================================
Private Sub EnsureSheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
        ws.Tab.Color = RGB(148, 163, 184)
    End If
End Sub

Private Function NextRunId() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(HIST_FAC_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        NextRunId = 1
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then
        NextRunId = 1
        Exit Function
    End If
    Dim maxId As Long : maxId = 0
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 2).Value) Then
            If CLng(ws.Cells(r, 2).Value) > maxId Then
                maxId = CLng(ws.Cells(r, 2).Value)
            End If
        End If
    Next r
    NextRunId = maxId + 1
End Function

Private Function FormatDiff(v As Double, fmt As String) As String
    If v > 0 Then
        FormatDiff = "+" & Format(v, fmt)
    ElseIf v < 0 Then
        FormatDiff = Format(v, fmt)
    Else
        FormatDiff = "±0"
    End If
End Function

Private Function DiffColor(v As Double) As Long
    If v > 0 Then
        DiffColor = RGB(231, 76, 60)    ' 赤: 増加
    ElseIf v < 0 Then
        DiffColor = RGB(41, 128, 185)   ' 青: 減少
    Else
        DiffColor = RGB(39, 174, 96)    ' 緑: 変化なし
    End If
End Function
