Sub DeleteClose()

    Dim ws  As Worksheet
    Dim r   As Long
    Dim idx As Long

    Set ws = ThisWorkbook.Sheets("議事録")

    ' CLOSE行を下から削除
    For r = 185 To 6 Step -1
        If Trim(CStr(ws.Cells(r, 5).Value)) = "CLOSE" Then
            ws.Rows(r).Delete Shift:=xlUp
        End If
    Next r

    ' A列に連番を振り直す（空白行はスキップ）
    idx = 1
    For r = 6 To 185
        If ws.Cells(r, 2).Value <> "" Then
            ws.Cells(r, 1).Value = idx
            idx = idx + 1
        Else
            ws.Cells(r, 1).Value = ""
        End If
    Next r

    ' 罫線を再設定
    With ws.Range("A6:F185").Borders
        .LineStyle = xlContinuous
        .Weight    = xlThin
        .Color     = RGB(144, 164, 174)
    End With

    MsgBox "CLOSE行を削除し、連番を振り直しました。", vbInformation

End Sub
