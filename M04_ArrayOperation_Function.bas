Attribute VB_Name = "M04_ArrayOperation_Function"
'Option Explicit
Dim tempArr()

'���I�ɔz��ɐV���ȗv�f��������
Function ArrayAdd(argArr() As Variant, ByVal data As Variant) As Variant()
    On Error GoTo err
    ReDim Preserve argArr(UBound(argArr) + 1): GoTo endProc
err:
    ReDim argArr(0): GoTo endProc
endProc:
    argArr(UBound(argArr)) = data
    ArrayAdd = argArr
End Function

'�z��̍Ō�̗v�f���폜����
Function ArrayCut(argArr() As Variant) As Variant()
    On Error Resume Next
    ReDim Preserve argArr(UBound(argArr) - 1)
    ArrayCut = argArr
End Function

'Variant�^�̔z��̒��g�̕����S�Ĉ�v����Ƃ��A���̌^��Ԃ�
Function ArrayType(ByVal argArr As Variant) As String
    Dim typenmArr()
    For Each arr In argArr
        typenmArr = ArrayAdd(typenmArr, TypeName(arr))
    Next
    If UBound(Filter(typenmArr, typenmArr(0))) = UBound(typenmArr) Then
        ArrayType = typenmArr(0)
    Else
        ArrayType = "Nothing"
    End If
End Function



Function RangeToArray(ByVal argRng As Range) As Variant()   '�͈͂�񎟌��z��ɕϊ�
    ReDim tempArr(argRng.Rows.Count, argRng.Columns.Count)
    For r = 1 To argRng.Rows.Count
        For c = 1 To argRng.Columns.Count
            tempArr(r - 1, c - 1) = argRng(r, c).Value
        Next
    Next
    RangeToArray = tempArr
End Function

Function Matrix(argArr As Variant) As Variant()     '�s��u��
    ReDim tempArr(UBound(argArr, 2), UBound(argArr, 1))
    For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
        For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
            tempArr(c, r) = argArr(r, c)
        Next
    Next
    Matrix = tempArr
End Function

Function RowIntervalAdd(argArr As Variant, Optional interval As Long = 1) As Variant() 'interval�s���Ԋu������
    ReDim tempArr((UBound(argArr, 1) - 1) * (interval + 1) + 1, UBound(argArr, 2))
    For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
        For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
            tempArr(r * (interval + 1), c) = argArr(r, c)
        Next
    Next
    RowIntervalAdd = tempArr
End Function

Function ColIntervalAdd(argArr As Variant, Optional interval As Long = 1) As Variant() 'interval�񂸂Ԋu������
    ReDim tempArr(UBound(argArr, 1), (UBound(argArr, 2) - 1) * (interval + 1) + 1)
    For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
        For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
            tempArr(r, c * (interval + 1)) = argArr(r, c)
        Next
    Next
    ColIntervalAdd = tempArr
End Function

Sub ArrayPaste(argArr As Variant, ByVal destCell As Range)  '�z��͈͓̔\��t��
    For r = 0 To UBound(argArr, 1) - 1
        For c = 0 To UBound(argArr, 2) - 1
            If argArr(r, c) <> "" Then destCell.Parent.Cells(destCell.Row + r, destCell.Column + c).Value = argArr(r, c)
        Next
    Next
End Sub

'�z��ɔz��ǉ�(direction:1���E�A2�����A3�����A4����)
Function ArrayAddArray(argArr1() As Variant, argArr2() As Variant, Optional direction As Long = 1) As Variant()
    If direction >= 1 And direction <= 4 Then
        If direction = 1 Or direction = 3 Then ReDim tempArr(UBound(argArr1, 1), UBound(argArr1, 2) + UBound(argArr2, 2))
        If direction = 2 Or direction = 4 Then ReDim tempArr(UBound(argArr1, 1) + UBound(argArr2, 1), UBound(argArr1, 2))
        
        If direction = 1 Or direction = 2 Then firstArr = argArr1: secondArr = argArr2
        If direction = 3 Or direction = 4 Then firstArr = argArr2: secondArr = argArr1
        
        For r = LBound(firstArr, 1) To UBound(firstArr, 1) - 1
            For c = LBound(firstArr, 2) To UBound(firstArr, 2) - 1
                tempArr(r, c) = firstArr(r, c)
            Next
        Next

        For r = LBound(secondArr, 1) To UBound(secondArr, 1) - 1
            For c = LBound(secondArr, 2) To UBound(secondArr, 2) - 1
                If direction = 1 Or direction = 3 Then tempArr(r, UBound(firstArr, 2) + c) = secondArr(r, c)
                If direction = 2 Or direction = 4 Then tempArr(UBound(firstArr, 1) + r, c) = secondArr(r, c)
            Next
        Next
        ArrayAddArray = tempArr
    End If
End Function
