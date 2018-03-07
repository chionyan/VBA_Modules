Attribute VB_Name = "M04_ArrayOperation_Function"
'Option Explicit
Dim tempArr()

'2つの配列の共通部分のみ取り出す
Function DuplicationArray(ByVal arr1, ByVal arr2) As Variant()
    For Each Item1 In arr1
        For Each Item2 In arr2
            If Item1 = Item2 Then DuplicationArray = ArrayAdd(DuplicationArray, Item2)
        Next
    Next
End Function

'動的に配列に新たな要素を加える
Function ArrayAdd(argArr() As Variant, ByVal data As Variant) As Variant()
    On Error GoTo Err
    ReDim Preserve argArr(UBound(argArr) + 1): GoTo endProc
Err:
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

'配列内のargsの要素数を返す
Function ArrayCount(argArr, ByVal args As Variant) As Long
    On Error GoTo errProc
    ArrayCount = UBound(Filter(argArr, args)) + 1
    Exit Function
errProc:
    ArrayCount = 0
End Function

Function RangeToArray(ByVal argRng As Range) As Variant()   '範囲を二次元配列に変換
    If Not argRng Is Nothing Then
        ReDim tempArr(argRng.Rows.Count, argRng.Columns.Count)
        For r = 1 To argRng.Rows.Count
            For c = 1 To argRng.Columns.Count
                tempArr(r - 1, c - 1) = argRng(r, c).Value
            Next
        Next
        RangeToArray = tempArr
    End If
End Function

Function RangeToOneDimention(ByVal argRng As Range) As String() '範囲内の値を一次元配列化
    If Not argRng Is Nothing Then
        For Each area In argRng.Areas
            For Each rng In area
                tempStr = tempStr & rng.Value & ","
            Next
        Next
        tempStr = Mid(tempStr, 1, Len(tempStr) - 1)
        RangeToOneDimention = Split(tempStr, ",")
    End If
End Function

'配列内の重複削除
Function ArrayDeduplication(argArr As Variant) As Variant()
    Dim tempArr() As Variant
    Dim col As New Collection
    Dim i As Long

    For i = LBound(argArr) To UBound(argArr)
        On Error Resume Next
        col.Add argArr(i), CStr(argArr(i))
        If Err.Number = 0 Then
            ReDim Preserve tempArr(col.Count - 1)
            tempArr(col.Count - 1) = argArr(i)
        End If
        On Error GoTo 0
    Next
    Set col = Nothing
    ArrayDeduplication = tempArr

End Function

Function Matrix(argArr As Variant) As Variant()     '行列置換
    If IsArrayEx(argArr) = 1 Then
        ReDim tempArr(UBound(argArr, 2), UBound(argArr, 1))
        For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
            For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
                tempArr(c, r) = argArr(r, c)
            Next
        Next
        Matrix = tempArr
    End If
End Function

Function RowIntervalAdd(argArr As Variant, Optional interval As Long = 1) As Variant() 'interval行ずつ間隔あける
    If IsArrayEx(argArr) = 1 Then
        ReDim tempArr((UBound(argArr, 1) - 1) * (interval + 1) + 1, UBound(argArr, 2))
        For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
            For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
                tempArr(r * (interval + 1), c) = argArr(r, c)
            Next
        Next
        RowIntervalAdd = tempArr
    End If
End Function

Function ColIntervalAdd(argArr As Variant, Optional interval As Long = 1) As Variant() 'interval列ずつ間隔あける
    If IsArrayEx(argArr) = 1 Then
        ReDim tempArr(UBound(argArr, 1), (UBound(argArr, 2) - 1) * (interval + 1) + 1)
        For r = LBound(argArr, 1) To UBound(argArr, 1) - 1
            For c = LBound(argArr, 2) To UBound(argArr, 2) - 1
                tempArr(r, c * (interval + 1)) = argArr(r, c)
            Next
        Next
        ColIntervalAdd = tempArr
    End If
End Function

Sub ArrayPaste(argArr As Variant, ByVal tgtCell As Range)  '配列の範囲貼り付け
    If IsArrayEx(argArr) = 1 Then
        For r = 0 To UBound(argArr, 1) - 1
            For c = 0 To UBound(argArr, 2) - 1
                If argArr(r, c) <> "" Then tgtCell.Parent.Cells(tgtCell.Row + r, tgtCell.Column + c).Value = argArr(r, c)
            Next
        Next
    End If
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

