Sub SortAgenda()

    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim rng     As Range
    Dim r       As Long
    Dim idx     As Long

    Set ws = ThisWorkbook.Sheets("議事録")

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 6 Then
        MsgBox "並び替えるデータがありません。", vbExclamation
        Exit Sub
    End If

    Set rng = ws.Range("B6:F" & lastRow)

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("B6:B" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, _
        CustomOrder:="新規,継続", _
        DataOption:=xlSortNormal

    With ws.Sort
        .SetRange rng
        .Header      = xlNo
        .MatchCase   = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' 連番を振り直す
    idx = 1
    For r = 6 To 185
        If ws.Cells(r, 2).Value <> "" Then
            ws.Cells(r, 1).Value = idx
            idx = idx + 1
        End If
    Next r

    MsgBox "種類順に並び替えました。", vbInformation

End Sub
