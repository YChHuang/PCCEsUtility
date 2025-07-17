## 將PCCES的溢出的列合併
PCCES生成的標單會自動把裝不下的文字溢出到下一列，這時可以用這個vba解決，目前只會合併B欄，有需求請自行擴充。

使用方法:  
* ALT+F11打開VBA介面
* 在左方的專案Project找到要新增的檔案->模組->插入->模組
* 貼上VBA就可以執行了
* 目前是使用ActiveSheet也就是使用中工作表，也就是目前用哪個檔案就用它來執行
* 首先會詢問要資料的起始列，這避免操作到一些標題之類的
* 判斷邏輯是"項次為空且沒有單位 = 溢出列"符合條件的標單請不要使用
* 請注意不要有小計、備註等奇怪的列
* VBA會跑比較久請耐心等待

原本溢出狀態  
<img width="947" height="193" alt="螢幕擷取畫面 2025-07-17 150851" src="https://github.com/user-attachments/assets/f7eae233-6160-4cbe-895b-1c131d48f70f" />  
執行後  
<img width="946" height="147" alt="螢幕擷取畫面 2025-07-17 150919" src="https://github.com/user-attachments/assets/68670391-0d7a-476c-9c59-f55b66f38bcf" />  


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
