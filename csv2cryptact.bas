Attribute VB_Name = "Module1"
Option Explicit

Sub XPC_csv_to_cryptact()
    
'シートを挿入する
    Worksheets.Add.Name = "XPC qtWallet"
    
'実行結果"TRUE"のみ抽出
'自分への送金の行を抽出
'鋳造の行を抽出
    Worksheets(2).Select
    Range("A1").AutoFilter _
        Field:=1, Criteria1:="TRUE"
    Range("C1").AutoFilter _
        Field:=3, Criteria1:="鋳造", _
        Operator:=xlOr, Criteria2:="自分への送金"

'XPC qtWalletにコピーする
    Range("A2").CurrentRegion. _
        SpecialCells(xlCellTypeVisible).Copy (Worksheets("XPC qtWallet").Range("A1"))
    
'オートフィルタを解除
    Range("C1").AutoFilter

'列の並び替え
    Worksheets("XPC qtWallet").Select
    Columns(2).Copy 'Timestamp
    Columns(1).PasteSpecial
    Columns(3).Copy 'Action
    Columns(2).PasteSpecial
    Columns(5).Copy 'Source
    Columns(3).PasteSpecial
    Columns(7).Copy 'Comment
    Columns(12).PasteSpecial
    Columns(6).Copy 'Volume
    Columns(7).PasteSpecial
    Application.CutCopyMode = False

'最終行を取得
    Dim MaxRow As Long
    MaxRow = Cells(Rows.Count, "A").End(xlUp).Row

'列ごとに必要なデータを入力
    Range(Cells(2, 4), Cells(MaxRow, 4)).Value = "XPC" 'Base
    Range(Cells(2, 9), Cells(MaxRow, 9)).Value = "JPY" 'Counter
    Range(Cells(2, 10), Cells(MaxRow, 10)).Value = "0" 'Fee
    Range(Cells(2, 11), Cells(MaxRow, 11)).Value = "JPY" 'FeeCcy
    Range(Cells(2, 5), Cells(MaxRow, 6)).Value = "" 'DerivType
    Range(Cells(2, 8), Cells(MaxRow, 8)).Value = "" 'DerivDetails

'TimestampをCryptact仕様に変更
    Range(Cells(2, 1), Cells(MaxRow, 1)).Replace what:="-", replacement:="/"
    Range(Cells(2, 1), Cells(MaxRow, 1)).Replace what:="T", replacement:=" "
    Range(Cells(2, 1), Cells(MaxRow, 1)).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"

'ActionをCryptact仕様に変更
    Range(Cells(2, 2), Cells(MaxRow, 2)).Replace "鋳造", "MINING"
    Range(Cells(2, 2), Cells(MaxRow, 2)).Replace "自分への送金", "SENDFEE"

'1行目に項目を入力
    Range("A1").Value = "Timestamp"
    Range("B1").Value = "Action"
    Range("C1").Value = "Source"
    Range("D1").Value = "Base"
    Range("E1").Value = "DerivType"
    Range("F1").Value = "DerivDetails"
    Range("G1").Value = "Volume"
    Range("H1").Value = "Price"
    Range("I1").Value = "Counter"
    Range("J1").Value = "Fee"
    Range("K1").Value = "FeeCcy"
    Range("L1").Value = "Comment"

End Sub
