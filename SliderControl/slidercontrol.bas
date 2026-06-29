Option Explicit

Sub MoveSuiiPreMonth()

    DisableUpdating
    
    
    Dim tblSheetWork As ListObject
    Set tblSheetWork = wsSheetWork.ListObjects("Sheet1テーブル")

    Dim Sheet() As Variant
    Sheet = tblSheetWork.DataBodyRange.value

    Dim dtMin As Date
    dtMin = Sheet(LBound(Sheet), SheetIndex.Hiduke)
    dtMin = DateSerial(year(dtMin), month(dtMin), 1)
    
    Dim dtStart As Date
    dtStart = Sheet2.Range("B1").value

    If dtMin < dtStart Then
        MoveSuiiMonth -1
    Else
        MsgBox "これより過去のデータがありません"
    End If
    
    
    EnableUpdating
    
End Sub

Sub MoveSuiiNextMonth()

    DisableUpdating
    
    
    Dim tblSheetWork As ListObject
    Set tblSheetWork = wsSheetWork.ListObjects("Sheet1テーブル")

    Dim Sheet() As Variant
    Sheet = tblSheetWork.DataBodyRange.value

    Dim dtMax As Date
    dtMax = Sheet(UBound(Sheet), SheetIndex.Hiduke)
    dtMax = DateSerial(year(dtMax), month(dtMax) + 1, 0)
    
    Dim dtEnd As Date
    dtEnd = Sheet2.Range("M1").value

    If dtEnd < dtMax Then
        MoveSuiiMonth 1
    Else
        MsgBox "これより未来のデータがありません"
    End If


    EnableUpdating
    
End Sub

