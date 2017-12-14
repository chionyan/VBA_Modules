Attribute VB_Name = "M04_ArrayOperation_Function"
'Option Explicit
Dim tempArr()

'動的に配列に新たな要素を加える
Function ArrayAdd(argArr() As Variant, ByVal data As Variant) As Variant()
    On Error GoTo err
    ReDim Preserve argArr(UBound(argArr) + 1): GoTo endProc
err:
    ReDim argArr(0): GoTo endProc
endProc:
    argArr(UBound(argArr)) = data
    ArrayAdd = argArr
End Function

'配列の最後の要素を削除する
Function ArrayCut(argArr() As Variant) As Variant()
    On Error Resume Next
    ReDim Preserve argArr(UBound(argArr) - 1)
    ArrayCut = argArr
End Function

'Variant型の配列の中身の方が全て一致するとき、その型を返す
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



Function RangeToArray(ByVal argRng As Range) As Variant()   '範囲を二次元配列に変換
    ReDim tempArr(argRng.Rows.Count, argRng.Columns.Count)
    For r = 1 To argRng.Rows.Count
        For c = 1 To argRng.Columns.Count
            tempArr(r - 1, c - 1) = argRng(r, c).Value
        Next
    Next
    RangeToArray = tempArr
End Function

Function Matrix(argArr As Variant) As Variant()     '行列置換
    ReDim tempArr(UBound(argArr, 2), UBound(argArr, 1))
    For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
        For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
            tempArr(c, r) = argArr(r, c)
        Next
    Next
    Matrix = tempArr
End Function

Function RowIntervalAdd(argArr As Variant, Optional interval As Long = 1) As Variant() 'interval行ずつ間隔あける
    ReDim tempArr((UBound(argArr, 1) - 1) * (interval + 1) + 1, UBound(argArr, 2))
    For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
        For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
            tempArr(r * (interval + 1), c) = argArr(r, c)
        Next
    Next
    RowIntervalAdd = tempArr
End Function

Function ColIntervalAdd(argArr As Variant, Optional interval As Long = 1) As Variant() 'interval列ずつ間隔あける
    ReDim tempArr(UBound(argArr, 1), (UBound(argArr, 2) - 1) * (interval + 1) + 1)
    For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
        For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
            tempArr(r, c * (interval + 1)) = argArr(r, c)
        Next
    Next
    ColIntervalAdd = tempArr
End Function

Sub ArrayPaste(argArr As Variant, ByVal destCell As Range)  '配列の範囲貼り付け
    For r = 0 To UBound(argArr, 1) - 1
        For c = 0 To UBound(argArr, 2) - 1
            If argArr(r, c) <> "" Then destCell.Parent.Cells(destCell.Row + r, destCell.Column + c).Value = argArr(r, c)
        Next
    Next
End Sub

'配列に配列追加(direction:1→右、2→下、3→左、4→上)
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
