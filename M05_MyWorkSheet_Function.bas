Attribute VB_Name = "M05_MyWorkSheet_Function"
'******************************************************************************************
' Procedure名：CONCAT(argRng As Range) As String
' 機能概要　  ：選択範囲の文字列を結合した値を返す関数
'******************************************************************************************
Function CONCAT(ByVal argRng As Range) As String
    Application.Volatile
    For Each rng In argRng
        CONCAT = CONCAT + rng.Value
    Next rng
End Function

'******************************************************************************************
' Procedure名：MAXROW(argRng As Range) As Long
' 機能概要　  ：選択範囲の一番セル行数を取得する関数 (空白有にも対応)
'******************************************************************************************
Function MAXROW(ByVal argRng As Range) As Long
    Application.Volatile
    For Each area In argRng.Areas
        Dim startCell As Range: Set startCell = area(1)
        Dim lastCell As Range: Set lastCell = area(area.Count)
        If lastCell.Row <> Rows.Count Then Set lastCell = lastCell.Offset(1, 0)
        For Column = startCell.Column To lastCell.Column
            maxrow_temp = Sheets(argRng.Parent.Name).Cells(lastCell.Row, Column).End(xlUp).Row
            If maxrow_temp >= MAXROW Then MAXROW = maxrow_temp
        Next
    Next
End Function
