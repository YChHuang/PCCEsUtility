## 將PCCES的溢出的列合併
PCCES生成的標單會自動把裝不下的文字溢出到下一列，這時可以用這個vba解決，目前只會合併B欄，有需求請自行擴充。

使用方法:  
* 請務必備份檔案或資料後再執行
* ALT+F11打開VBA介面
* 在左方的專案Project找到要新增的檔案->模組->插入->模組
* 貼上VBA就可以執行了

### 效果

![死圖](./VBA_Utility_SampleFiles/sample.png)
![死圖](./VBA_Utility_SampleFiles/origin.png)
![死圖](./VBA_Utility_SampleFiles/MergePreProcess.png)
![死圖](./VBA_Utility_SampleFiles/MergeResult.png)


```VBA
Sub MergePCCSformWithoutTitle_ByUnit()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim StartIndex As Long
    Dim tmp As Long
    Dim dummy As Range
    
    ' ==========README===========================================================
    '
    ' 把 Pcces 生成的 Excel/CSV 檔案中溢出到下一列的內容合併
    ' 資料過大要等他跑完，不然會壞掉
    ' 一行行刪除雖然很白癡，但中斷後隨時可以保存進度，因為vba太爛了，這樣可靠性最高
    ' 判斷邏輯是這樣:
    ' A欄有數值且C欄有數值，則此列為原始資料列。
    ' 若非上述之列，則判定為溢出，向上合併後刪除。
    ' 使用前請先檢查結構是不是這樣。
    '
    ' ============================================================================


    ' 設定目前工作表為當前開啟之工作簿使用中之工作表
    
    Set ws = ActiveWorkbook.ActiveSheet
    
    
    
    ' 取得 B 欄最後一列的行號
    '找到B欄位最底層(無視所有資料)後向上查找到第一個有值的欄位，回傳它的所在row
    
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    '找到第一列的最右欄儲存格(無視所有資料)後，向左查找，回傳它的所在column
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
    
    '尋找標題列並略過，注意一下，如果撞上合併儲存格會報錯
    StartIndex = InputBox("資料從哪列開始?")
    
    ' 由下向上遍歷data，因為刪除儲存格操作會由下方列遞補當前index，若由上到下遍歷，會導致略過。
                    
    For i = lastRow To StartIndex Step -1
        ' 更新狀態列，顯示當前進度
        Application.StatusBar = "正在處理行: " & lastRow - i & " / " & lastRow
        
        '如果A欄為空並且C欄為空，則判定為溢出行
        If ws.Cells(i, 1).Value = "" And ws.Cells(i, 3).Value = "" Then
            '新值 = 舊值 & 溢出欄位之值
            ws.Cells(i - 1, 2).Value = ws.Cells(i - 1, 2).Value & ws.Cells(i, 2)
            ws.Rows(i).Delete
            '溢出欄位整列刪除
        End If
    Next i
    
        
    
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
