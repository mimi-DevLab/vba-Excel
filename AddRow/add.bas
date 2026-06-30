テーブルに新規行を追加するボタン

Public Sub AddRow()

    Dim tbl As ListObject

    Set tbl = wsSettings.ListObjects("テーブル")

    tbl.ListRows.Add

End Sub