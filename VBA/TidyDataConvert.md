## 將PCCES的標單生成的項次依照政則化邏輯把標題用標籤方式貼在資料後方
用正則化分析出大中小標題，並將他們整理成TidyData的方式貼到每一列，以方編作樞紐分析表或sumifs等用途。

使用方法:  
* ALT+F11打開VBA介面
* 在左方的專案Project找到要新增的檔案->模組->插入->模組
* 貼上VBA就可以執行了
* 目前是使用ActiveSheet也就是使用中工作表，也就是目前用哪個檔案就用它來執行
* 首先會詢問要將標籤貼在哪一欄之後
* 再來會遍例A欄找到符合標題的正則化條件(可以依照需求修改)
* 填入後，會在執行一次向下填滿


  
```VBA
Option Explicit

Sub TidyDataLabels()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim inOff As Variant
    Dim off As Long
    Dim resp As VbMsgBoxResult
    
    '─── 1. 讀取並驗證 offset ─────────────────────────
    Do
        inOff = Application.InputBox( _
            Prompt:="請輸入輸出欄位起始 offset（整數，從第幾欄開始向右寫入）：", _
            Title:="設定 offset", Type:=1)
        If inOff = False Then Exit Sub    ' 使用者按「取消」
        If IsNumeric(inOff) And inOff >= 1 Then
            off = CLng(inOff)
            Exit Do
        Else
            MsgBox "請輸入正整數！", vbExclamation
        End If
    Loop
    
    resp = MsgBox( _
        "接下來的資料將從第 " & off & " 欄開始向右覆寫，請確認此範圍沒有重要資料！", _
        vbExclamation + vbOKCancel, "覆寫警告")
    If resp = vbCancel Then Exit Sub
    
    '─── 2. 建立三個正則物件 ────────────────────────────
    Dim re1 As Object, re2 As Object, re3 As Object, re4 As Object
    Set re1 = CreateObject("VBScript.RegExp")
    Set re2 = CreateObject("VBScript.RegExp")
    Set re3 = CreateObject("VBScript.RegExp")
    Set re4 = CreateObject("VBScript.RegExp")
    
    With re1
        .Pattern = "^[甲癸酉子丑]\.[壹貳參肆伍]\.[一二三四五六七八九十]+$"
        .Global = False
    End With
    With re2
        .Pattern = "^[甲癸酉子丑]\.[壹貳參肆伍]\.[一二三四五六七八九十]+\.\d+$"
        .Global = False
    End With
    With re3
        .Pattern = "^[甲癸酉子丑]\.[壹貳參肆伍]\.[一二三四五六七八九十]+\.\d+\.\d+$"
        .Global = False
    End With
    With re4
        .Pattern = "^[甲癸酉子丑]\.[壹貳參肆伍]\.[一二三四五六七八九十]+\.\d+\.\d+\.\d+$"
        .Global = False
    End With
    
    '─── 3. 掃描 A 欄，將符合層級的原文貼到 offset+n ─────────
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Dim i As Long
    Dim txt As String
    Dim content As String
    
    For i = 2 To lastRow
        txt = CStr(ws.Cells(i, "A").Value)
        content = ws.Cells(i, 2).Value
        
        If re1.test(txt) Then
            ws.Cells(i, off + 1).Value = content
        End If
        If re2.test(txt) Then
            ws.Cells(i, off + 2).Value = content
        End If
        If re3.test(txt) Then
            ws.Cells(i, off + 3).Value = content
        End If
        If re4.test(txt) Then
            ws.Cells(i, off + 4).Value = content
        End If
    Next i
    
    '─── 4. 各欄向下填滿空白 ───────────────────────────
    Dim c As Long
    For c = off + 1 To off + 4
        For i = 3 To lastRow
            If ws.Cells(i, c).Value = "" Then
                ws.Cells(i, c).Value = ws.Cells(i - 1, c).Value
            End If
        Next i
    Next c
    
    MsgBox "Tidy data 標籤處理完成！", vbInformation
End Sub



```
