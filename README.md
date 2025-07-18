# PCCEsUtility
常常處理PCCES轉Excel的標單時，常常遇到儲存格溢出到下一列的問題，以至於要篩選、資料分析都很困難，  
因此我用AI協助整理了一些VBA希望可以幫助大家解決這個問題。

## 功能特色
* 合併PCCES轉成Excel檔案時的溢出欄位
* 幫大家省下一行一行剪下貼上的時間
* 方便日後做成樞紐分析表或是查找資料時更方便

## 目前我整理兩個常用的功能
* [合併溢出列](https://github.com/YChHuang/PCCEsUtility/blob/main/VBA/MergePCCSformWithoutTitle.md)
* [將標單轉成TidyData](https://github.com/YChHuang/PCCEsUtility/blob/main/VBA/TidyDataConvert.md)

## 使用方法
* 請務必備份檔案或資料後再執行
* 開啟Excel
* 允許Excel執行巨集
* ALT+F11打開VBA介面
* 在左方的專案Project找到要新增的檔案->模組->插入->模組
* 貼上VBA就可以執行了
