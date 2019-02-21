Attribute VB_Name = "Module1"
Option Explicit

Sub XPC_csv_to_cryptact()
    
'�V�[�g��}������
    Worksheets.Add.Name = "XPC qtWallet"
    
'���s����"TRUE"�̂ݒ��o
'�����ւ̑����̍s�𒊏o
'�����̍s�𒊏o
    Worksheets(2).Select
    Range("A1").AutoFilter _
        Field:=1, Criteria1:="TRUE"
    Range("C1").AutoFilter _
        Field:=3, Criteria1:="����", _
        Operator:=xlOr, Criteria2:="�����ւ̑���"

'XPC qtWallet�ɃR�s�[����
    Range("A2").CurrentRegion. _
        SpecialCells(xlCellTypeVisible).Copy (Worksheets("XPC qtWallet").Range("A1"))
    
'�I�[�g�t�B���^������
    Range("C1").AutoFilter

'��̕��ёւ�
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

'�ŏI�s���擾
    Dim MaxRow As Long
    MaxRow = Cells(Rows.Count, "A").End(xlUp).Row

'�񂲂ƂɕK�v�ȃf�[�^�����
    Range(Cells(2, 4), Cells(MaxRow, 4)).Value = "XPC" 'Base
    Range(Cells(2, 9), Cells(MaxRow, 9)).Value = "JPY" 'Counter
    Range(Cells(2, 10), Cells(MaxRow, 10)).Value = "0" 'Fee
    Range(Cells(2, 11), Cells(MaxRow, 11)).Value = "JPY" 'FeeCcy
    Range(Cells(2, 5), Cells(MaxRow, 6)).Value = "" 'DerivType
    Range(Cells(2, 8), Cells(MaxRow, 8)).Value = "" 'DerivDetails

'Timestamp��Cryptact�d�l�ɕύX
    Range(Cells(2, 1), Cells(MaxRow, 1)).Replace what:="-", replacement:="/"
    Range(Cells(2, 1), Cells(MaxRow, 1)).Replace what:="T", replacement:=" "
    Range(Cells(2, 1), Cells(MaxRow, 1)).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"

'Action��Cryptact�d�l�ɕύX
    Range(Cells(2, 2), Cells(MaxRow, 2)).Replace "����", "MINING"
    Range(Cells(2, 2), Cells(MaxRow, 2)).Replace "�����ւ̑���", "SENDFEE"

'1�s�ڂɍ��ڂ����
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
