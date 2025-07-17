這能夠把PCCES轉Excel後溢出的列給合併
```VBA
Sub MergePCCSformWithoutTitle_ByUnit()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim StartIndex As Long
    Dim tmp As Long
    Dim dummy As Range
    
    ' 把 Pcces 生成的 Excel 檔案中溢出到下一列的內容合併
    '___V.1.2.0___

    ' 設定目前工作表
    Set ws = ActiveWorkbook.ActiveSheet
    
    ' 取得 B 欄最後一列的行號
    '找到B欄位最底層(無視所有資料)後向上查找到第一個有值的欄位，回傳它的所在row
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    '找到第一列的最右欄儲存格(無視所有資料)後，向左查找，回傳它的所在column
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
    '尋找標題列並略過，如果撞上合併儲存格會報錯
    StartIndex = InputBox("資料從哪列開始?")
    
    ' 由下向上遍歷data，因為刪除儲存格操作會由下方列遞補當前index，若由上到下遍歷，會導致略過。
    
    For i = lastRow To StartIndex Step -1
        ' 更新狀態列，顯示當前進度
        Application.StatusBar = "正在處理行: " & lastRow - i & " / " & lastRow
        
        '如果A欄為空並且C欄為空，則判定為溢出行
        If ws.Cells(i, 1).Value = "" And ws.Cells(i, 3).Value = "" Then
            '新值 = 舊值 & 溢出欄位之值
            ws.Cells(i - 1, 2).Value = ws.Cells(i - 1, 2).Value & ws.Cells(i, 2)
            '溢出欄位整列刪除
        End If
    Next i
    
    '用autofilter刪除空格
    With ws.Range("A" & StartIndex & ":C" & lastRow)
        .AutoFilter Field:=1, Criteria1:="="
        .AutoFilter Field:=3, Criteria1:="="
        
        On Error Resume Next
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        ws.AutoFilterMode = False
    End With
        
    
    'B欄設定自動換列
    ws.Columns(2).WrapText = True

    ' 還原狀態列
    

    '重設scrollBar，可以註解看看少了這段會怎樣，把右邊的scrollBar下拉到底
    '強制要求excel計算UsedRange，超低能但只要呼叫他就重設了
    Set dummy = ws.UsedRange
    dummy.EntireRow.AutoFit
    
    
    Application.StatusBar = False
    MsgBox "已將溢出儲存格合併", vbInformation, "完成"
End Sub

```
