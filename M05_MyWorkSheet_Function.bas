Attribute VB_Name = "M05_MyWorkSheet_Function"
'******************************************************************************************
' Procedure���FCONCAT(argRng As Range) As String
' �@�\�T�v�@  �F�I��͈͂̕���������������l��Ԃ��֐�
'******************************************************************************************
Function CONCAT(ByVal argRng As Range) As String
    Application.Volatile
    For Each rng In argRng
        CONCAT = CONCAT + rng.Value
    Next rng
End Function

'******************************************************************************************
' Procedure���FMAXROW(argRng As Range) As Long
' �@�\�T�v�@  �F�I��͈͂̈�ԃZ���s�����擾����֐� (�󔒗L�ɂ��Ή�)
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
