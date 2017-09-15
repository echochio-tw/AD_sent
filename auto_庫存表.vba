Sub auto1()
    ' 刪除已存在的總倉
    For i = 2 To Sheets.Count
        If "總倉" = Sheets(i).Name Then
            Application.DisplayAlerts = False
            Sheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next
    
    ' A4:A5 + shift-ctrl + 下 key + 右 key
    With Sheets("Page1")
        Set rData = .Range(.Range("A4"), .Range("A5").End(xlToRight))
        Set rData = .Range(rData, rData.End(xlDown))
    End With
    
    '樞紐分析暫存
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "樞紐分析暫存"
    
    '樞紐分析建立
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        rData, Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="樞紐分析暫存!R3C1", TableName:="PivotTable1", DefaultVersion:= _
        xlPivotTableVersion10
        
    '樞紐分析欄位設定
    Sheets("樞紐分析暫存").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("產品編號")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("品名規格")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("類別名稱")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("倉庫名稱")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    
    '樞紐分析欄位設定加總
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables("PivotTable1" _
        ).PivotFields("實際在庫存量"), "加總 - 實際在庫存量", xlSum
        
    '樞紐分析欄位取消為無
    ActiveSheet.PivotTables("PivotTable1").PivotFields("產品編號").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("品名規格").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("類別名稱").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    Cells.Select
    
    '樞紐分析暫存複製到總倉, 複製值
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "總倉"
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.EntireColumn.AutoFit
    
    '選範圍
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    ' 將選的範圍 "0" 變 空白
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
         SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
         ReplaceFormat:=False
    
    ' 刪除樞紐分析暫存
    Application.DisplayAlerts = False
    Sheets("樞紐分析暫存").Delete
    Application.DisplayAlerts = True
    
    ' 選取 總倉
    Worksheets("總倉").Activate
    
    ' 找出 Row 及 Column 大小
    lc = Cells(4, 1).End(xlToRight).Column
    lr = Cells(4, 1).End(xlDown).Row
    
    ' 最後一行(Row)的所有 上的資訊為 0時刪除 其Column
    ' 要重後面開始刪除不然會算錯
    For xlc = lc To 4 Step -1
      If Cells(lr, xlc).Value = "0" Then
       ActiveSheet.Range(Cells(1, xlc), Cells(1, xlc)).EntireColumn.Delete
    Next xlc
    
End Sub
