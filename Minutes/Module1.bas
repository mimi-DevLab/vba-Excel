Option Explicit

' 議事録担当者リスト
Private Const NAMES_LIST As String = "Aさん,Bさん,Cさん,Dさん"

' データ範囲・列番号定数
Private Const DATA_START_ROW As Long = 6
Private Const DATA_END_ROW   As Long = 185
Private Const COL_TYPE       As Long = 2  ' B列: 種類
Private Const COL_TITLE      As Long = 3  ' C列: 議題名
Private Const COL_DETAIL     As Long = 4  ' D列: 詳細内容
Private Const COL_STATUS     As Long = 5  ' E列: 状況
Private Const COL_NOTE       As Long = 6  ' F列: 備考

' ---- [議事録作成] ボタン ---- 今日話し合った内容を整理してまとめてくれる
Sub Opening()

    ' ===== 変数宣言 =====
    Dim wsFormat      As Worksheet
    Dim wsTemplate    As Worksheet
    Dim wsNew         As Worksheet
    Dim newSheetName  As String
    Dim lastRow       As Long
    Dim k             As Long
    Dim i             As Long
    Dim j             As Long
    Dim tmpCol        As Long
    Dim tmp           As Variant
    Dim agendaList()  As Variant
    Dim oi            As Long
    Dim oj            As Long
    Dim doSwap        As Boolean
    Dim pasteRow      As Long
    Dim lastType      As String
    Dim lastName      As String
    Dim startRowA     As Long
    Dim startRowB     As Long
    Dim r             As Long
    Dim namesList()   As String
    Dim currentPerson As String
    Dim currentIdx    As Long
    Dim n             As Long
    Dim curDate       As Date
    Dim nextThu       As Date
    Dim daysToThu     As Long

    Set wsFormat   = ThisWorkbook.Sheets("議事録")
    Set wsTemplate = ThisWorkbook.Sheets("テンプレート")

    ' ----- 新規シート名（yyyymmdd_議事録）-----
    newSheetName = Format(Date, "yyyymmdd") & "_議事録"

    ' 同名チェック
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(newSheetName)
    On Error GoTo 0
    If Not wsNew Is Nothing Then
        MsgBox "「" & newSheetName & "」シートが既に存在します。", vbExclamation
        Exit Sub
    End If

    ' ----- データ収集 -----
    lastRow = wsFormat.Cells(wsFormat.Rows.Count, COL_TYPE).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "議題がありません。B列（種類）を入力してください。", vbExclamation
        Exit Sub
    End If

    ReDim agendaList(1 To (lastRow - DATA_START_ROW + 1), 1 To 4)
    k = 0

    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsFormat.Cells(i, COL_TYPE).Value)) <> "" Then
            k = k + 1
            agendaList(k, 1) = wsFormat.Cells(i, COL_TYPE).Value
            agendaList(k, 2) = wsFormat.Cells(i, COL_TITLE).Value
            agendaList(k, 3) = wsFormat.Cells(i, COL_DETAIL).Value
            agendaList(k, 4) = wsFormat.Cells(i, COL_NOTE).Value
        End If
    Next i

    If k = 0 Then
        MsgBox "有効な議題がありません。", vbExclamation
        Exit Sub
    End If

    ' ----- バブルソート（種類 → 議題名の2キー）-----
    For i = 1 To k - 1
        For j = i + 1 To k
            oi     = OrderValue(agendaList(i, 1))
            oj     = OrderValue(agendaList(j, 1))
            doSwap = False
            If oi > oj Then
                doSwap = True
            ElseIf oi = oj And CStr(agendaList(i, 2)) > CStr(agendaList(j, 2)) Then
                doSwap = True
            End If
            If doSwap Then
                For tmpCol = 1 To 4
                    tmp               = agendaList(i, tmpCol)
                    agendaList(i, tmpCol) = agendaList(j, tmpCol)
                    agendaList(j, tmpCol) = tmp
                Next tmpCol
            End If
        Next j
    Next i

    ' ----- テンプレートから新規シート作成 -----
    wsTemplate.Visible = xlSheetVisible
    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsTemplate.Visible = xlSheetHidden

    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsNew.Name = newSheetName

    ' メタ情報コピー
    wsNew.Range("B2").Value = wsFormat.Range("B2").Value
    wsNew.Range("B3").Value = wsFormat.Range("B3").Value
    wsNew.Range("B4").Value = wsFormat.Range("B4").Value

    ' ----- 出力処理 -----
    pasteRow  = DATA_START_ROW
    lastType  = ""
    lastName  = ""
    startRowA = pasteRow
    startRowB = pasteRow

    For i = 1 To k
        ' 種類が変わったら → A列グループ結合 & 新カテゴリ開始
        If CStr(agendaList(i, 1)) <> lastType Then
            If i > 1 Then
                MergeIfNeeded wsNew, startRowA, pasteRow - 1, 1
            End If
            wsNew.Cells(pasteRow, 1).Value = agendaList(i, 1)
            ApplyTypeColor wsNew, pasteRow, CStr(agendaList(i, 1))
            lastType  = CStr(agendaList(i, 1))
            startRowA = pasteRow
            lastName  = ""
            startRowB = pasteRow
        End If

        ' 議題名が変わったら → B列グループ結合 & 新議題開始
        If CStr(agendaList(i, 2)) <> lastName Then
            If i > 1 And lastName <> "" Then
                MergeIfNeeded wsNew, startRowB, pasteRow - 1, 2
            End If
            wsNew.Cells(pasteRow, 2).Value = agendaList(i, 2)
            lastName  = CStr(agendaList(i, 2))
            startRowB = pasteRow
        End If

        ' 詳細内容（C列）・備考（D列）
        wsNew.Cells(pasteRow, 3).Value    = agendaList(i, 3)
        wsNew.Cells(pasteRow, 3).WrapText = True
        wsNew.Cells(pasteRow, 4).Value    = agendaList(i, 4)

        pasteRow = pasteRow + 1
    Next i

    ' 最後のグループ結合
    MergeIfNeeded wsNew, startRowA, pasteRow - 1, 1
    MergeIfNeeded wsNew, startRowB, pasteRow - 1, 2

    ' ----- 「新規」→「継続」に更新 -----
    For r = DATA_START_ROW To DATA_END_ROW
        If Trim(CStr(wsFormat.Cells(r, COL_TYPE).Value)) = "新規" Then
            wsFormat.Cells(r, COL_TYPE).Value = "継続"
        End If
    Next r

    ' ----- 担当者ローテーション -----
    namesList     = Split(NAMES_LIST, ",")
    currentPerson = Trim(CStr(wsFormat.Range("B4").Value))
    currentIdx    = -1

    For n = 0 To UBound(namesList)
        If Trim(namesList(n)) = currentPerson Then
            currentIdx = n
            Exit For
        End If
    Next n

    If currentIdx = -1 Or currentIdx >= UBound(namesList) Then
        wsFormat.Range("B4").Value = Trim(namesList(0))
    Else
        wsFormat.Range("B4").Value = Trim(namesList(currentIdx + 1))
    End If

    ' ----- 次週木曜日を日付欄に設定 -----
    curDate   = Date
    daysToThu = 4 - Weekday(curDate, vbMonday)
    If daysToThu <= 0 Then daysToThu = daysToThu + 7
    nextThu = curDate + daysToThu
    wsFormat.Range("B2").Value = nextThu

    wsNew.Activate
    MsgBox "議事録シート「" & newSheetName & "」を作成しました。", vbInformation

End Sub

' 種類ごとの並び順（数値が小さいほど上に来る）
Private Function OrderValue(ByVal t As String) As Long
    Select Case Trim(t)
        Case "継続":     OrderValue = 1
        Case "新規":     OrderValue = 2
        Case "社内窓口": OrderValue = 3
        Case "備品確認": OrderValue = 4
        Case "保留":     OrderValue = 5
        Case Else:        OrderValue = 99
    End Select
End Function

' 必要なら複数行を結合する
Private Sub MergeIfNeeded(ByVal ws As Worksheet, _
                           ByVal startR As Long, _
                           ByVal endR As Long, _
                           ByVal col As Long)
    If endR > startR Then
        ws.Range(ws.Cells(startR, col), ws.Cells(endR, col)).Merge
        ws.Cells(startR, col).VerticalAlignment = xlCenter
    End If
End Sub

' 種類ごとに背景色を設定する
Private Sub ApplyTypeColor(ByVal ws As Worksheet, _
                            ByVal r As Long, _
                            ByVal typeName As String)
    Dim bg As Long
    Select Case Trim(typeName)
        Case "継続":     bg = RGB(232, 245, 233)
        Case "新規":     bg = RGB(227, 242, 253)
        Case "保留":     bg = RGB(255, 249, 196)
        Case "社内窓口": bg = RGB(252, 228, 236)
        Case "備品確認": bg = RGB(243, 229, 245)
        Case Else:        bg = RGB(255, 255, 255)
    End Select
    ws.Cells(r, 1).Interior.Color = bg
End Sub
