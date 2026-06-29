Option Explicit

'==============================
' 前月へ移動
'==============================
Sub MoveSuiiPreMonth()

    DisableUpdating

    ' Sheet1テーブルを配列へ読み込む
    Dim tbl As ListObject
    Set tbl = Sheet1.ListObjects("Sheet1テーブル")

    Dim arrData As Variant
    arrData = tbl.DataBodyRange.Value

    ' 最古の日付（月初へ丸める）
    Dim dtMin As Date
    dtMin = arrData(LBound(arrData), SheetIndex.Hiduke)
    dtMin = DateSerial(Year(dtMin), Month(dtMin), 1)

    ' 現在の表示開始日
    Dim dtStart As Date
    dtStart = Sheet2.Range("B1").Value

    ' まだ過去データがある場合のみ移動
    If dtMin < dtStart Then
        MoveSuiiMonth -1
    Else
        MsgBox "これより過去のデータがありません"
    End If

    EnableUpdating

End Sub


'==============================
' 翌月へ移動
'==============================
Sub MoveSuiiNextMonth()

    DisableUpdating

    ' Sheet1テーブルを配列へ読み込む
    Dim tbl As ListObject
    Set tbl = Sheet1.ListObjects("Sheet1テーブル")

    Dim arrData As Variant
    arrData = tbl.DataBodyRange.Value

    ' 最新の日付（月末へ丸める）
    Dim dtMax As Date
    dtMax = arrData(UBound(arrData), SheetIndex.Hiduke)
    dtMax = DateSerial(Year(dtMax), Month(dtMax) + 1, 0)

    ' 現在の表示終了日
    Dim dtEnd As Date
    dtEnd = Sheet2.Range("D1").Value

    ' まだ未来データがある場合のみ移動
    If dtEnd < dtMax Then
        MoveSuiiMonth 1
    Else
        MsgBox "これより未来のデータがありません"
    End If

    EnableUpdating

End Sub
