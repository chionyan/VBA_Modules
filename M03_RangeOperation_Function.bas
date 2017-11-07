Attribute VB_Name = "M03_RangeOperation_Function"
'******************************************************************************************
' Procedure名：FilterRng(ByVal argRng As Range, ByVal args As String) As Range
' 機能概要　  ：引数1の範囲内で引数2の正規表現に一致する範囲のみをフィルタリングして範囲を返す関数
'******************************************************************************************
Function FilterRng(ByVal argRng As Range, ByVal args As String) As Range
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = args
    Dim rng As Range
    For Each rng In argRng
        If re.Test(rng.Value) Then
            If FilterRng Is Nothing Then
                Set FilterRng = rng
            Else
                Set FilterRng = Union(FilterRng, rng)
            End If
        End If
    Next
End Function
'******************************************************************************************
' Procedure名：InputableRng(ByVal argRng As Range) As Range
' 機能概要　  ：引数1の範囲内で、編集可能なセルを返す関数
'******************************************************************************************
Function InputtableRng(ByVal argRng As Range) As Range
    Dim rng As Range
    For Each rng In argRng
        If rng.Locked = False Then
            If InputtableRng Is Nothing Then
                Set InputtableRng = rng
            Else
                Set InputtableRng = Union(InputRng, rng)
            End If
        End If
    Next rng
End Function
'******************************************************************************************
' Procedure名：BottomExtention(ByVal argRng As Range) As Range
' 機能概要　  ：引数1の範囲から一番下端の値の入ったセルまでの範囲を返す関数
'               ※現在引数一列のみ対応。今後複数列選択時にも対応予定。
'******************************************************************************************
Function BottomExtention(ByVal argRng As Range) As Range
    If argRng.Areas.Count = 1 Then
        Dim bottomRow As Long
        With Sheets(argRng.Parent.Name)
            bottomRow = .Cells(Rows.Count, argRng.Column).End(xlUp).Row
            Set BottomExtention = Range(argRng, .Cells(bottomRow, argRng.Column))
        End With
    End If
End Function
'******************************************************************************************
' Procedure名：RightExtention(ByVal argRng As Range) As Range
' 機能概要　  ：引数1の範囲から一番右端の値の入ったセルまでの範囲を返す関数
'               ※現在引数一行のみ対応。今後複数行選択時にも対応予定。
'******************************************************************************************
Function RightExtention(ByVal argRng As Range) As Range
    If argRng.Areas.Count = 1 Then
        Dim rightColumn As Long
        With Sheets(argRng.Parent.Name)
            rightColumn = .Cells(argRng.Row, Columns.Count).End(xlToLeft).Column
            Set RightExtention = Range(argRng, .Cells(argRng.Row, rightColumn))
        End With
    End If
End Function
'******************************************************************************************
' Procedure名：BottomRightExtention(ByVal argRng As Range) As Range
' 機能概要　  ：引数1の範囲から一番右下端の値の入ったセルまでの範囲を返す関数
'******************************************************************************************
Function BottomRightExtention(ByVal argRng As Range) As Range
    If argRng.Areas.Count = 1 Then
        Dim bottomRow As Long
        Dim rightColumn As Long
        With Sheets(argRng.Parent.Name)
            bottomRow = .Cells(Rows.Count, argRng.Column).End(xlUp).Row
            rightColumn = .Cells(argRng.Row, Columns.Count).End(xlToLeft).Column
            Set BottomRightExtention = Range(argRng, .Cells(bottomRow, rightColumn))
        End With
    End If
End Function
'******************************************************************************************
' Procedure名：Union2(ParamArray ArgList() As Variant) As Range
' 機能概要　  ：複数のセル ArgList の和集合を返す
'******************************************************************************************
Function Union2(ParamArray ArgList() As Variant) As Range
    Dim buf As Range
    Dim i As Long
    For i = 0 To UBound(ArgList)
        If TypeName(ArgList(i)) = "Range" Then
            If buf Is Nothing Then
                Set buf = ArgList(i)
            Else
                Set buf = Application.Union(buf, ArgList(i))
            End If
        End If
    Next
    Set Union2 = buf
End Function
'******************************************************************************************
' Procedure名：Intersect2(ParamArray ArgList() As Variant) As Range
' 機能概要　  ：複数のセル ArgList の積集合を返す
'******************************************************************************************
Function Intersect2(ParamArray ArgList() As Variant) As Range
    Dim buf As Range
    Dim i As Long
    For i = 0 To UBound(ArgList)
        If Not TypeName(ArgList(i)) = "Range" Then
            Exit Function
        ElseIf buf Is Nothing Then
            Set buf = ArgList(i)
        Else
            Set buf = Application.Intersect(buf, ArgList(i))
        End If
        If buf Is Nothing Then Exit Function
    Next
    Set Intersect2 = buf
End Function
'******************************************************************************************
' Procedure名：Except2(ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range
' 機能概要　  ：SourceRange から ArgList を差し引いた差集合を返す
'               (SourceRange と 反転した ArgList との積集合を返す)
'******************************************************************************************
Function Except2 _
    (ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range
    If TypeName(SourceRange) = "Range" Then
        Dim buf As Range
        Set buf = SourceRange
        Dim i As Long
        For i = 0 To UBound(ArgList)
            If TypeName(ArgList(i)) = "Range" Then
                Set buf = Intersect2(buf, Invert2(ArgList(i)))
            End If
        Next
        Set Except2 = buf
    End If
End Function
'******************************************************************************************
' Procedure名：Invert2(ByRef SourceRange As Variant) As Range
' 機能概要　  ：SourceRange の選択範囲を反転する
'*****************************************************************************************
Function Invert2(ByRef SourceRange As Variant) As Range
    If Not TypeName(SourceRange) = "Range" Then Exit Function
    Dim sh As Worksheet
    Set sh = SourceRange.Parent
    Dim buf As Range
    Set buf = SourceRange.Parent.Cells
    Dim a As Range
    For Each a In SourceRange.Areas
        Dim AreaTop    As Long
        Dim AreaBottom As Long
        Dim AreaLeft   As Long
        Dim AreaRight  As Long
        AreaTop = a.Row
        AreaBottom = AreaTop + a.Rows.Count - 1
        AreaLeft = a.Column
        AreaRight = AreaLeft + a.Columns.Count - 1
        
        '■□□
        '■×□
        '■□□  ■の部分
        Dim RangeLeft   As Range
        Set RangeLeft = GetRangeWithPosition(sh, _
            sh.Cells.Row, sh.Cells.Column, sh.Rows.Count, AreaLeft - 1)
        '   Top           Left             Bottom         Right

        '□□■
        '□×■
        '□□■  ■の部分
        Dim RangeRight  As Range
        Set RangeRight = GetRangeWithPosition(sh, _
            sh.Cells.Row, AreaRight + 1, sh.Rows.Count, sh.Columns.Count)
        '   Top           Left           Bottom         Right
        
        
        '□■□
        '□×□
        '□□□  ■の部分
        Dim RangeTop    As Range
        Set RangeTop = GetRangeWithPosition(sh, _
            sh.Cells.Row, AreaLeft, AreaTop - 1, AreaRight)
        '   Top           Left      Bottom       Right
        
        
        '□□□
        '□×□
        '□■□  ■の部分
        Dim RangeBottom As Range
        Set RangeBottom = GetRangeWithPosition(sh, _
            AreaBottom + 1, AreaLeft, sh.Rows.Count, AreaRight)
        '   Top              Left      Bottom         Right
        
        Set buf = Intersect2(buf, _
            Union2(RangeLeft, RangeRight, RangeTop, RangeBottom))
    Next
    Set Invert2 = buf
End Function
'******************************************************************************************
' Procedure名：Function GetRangeWithPosition(ByRef sh As Worksheet, ByVal Top As Long, ByVal Left As Long,
'                                               ByVal Bottom As Long, ByVal Right As Long) As Range
' 機能概要　  ：四隅を指定して Range を得る
'*****************************************************************************************
Function GetRangeWithPosition(ByRef sh As Worksheet, _
    ByVal Top As Long, ByVal Left As Long, _
    ByVal Bottom As Long, ByVal Right As Long) As Range
    
    '無効条件
    If Top > Bottom Or Left > Right Then
        Exit Function
    ElseIf Top < 0 Or Left < 0 Then
        Exit Function
    ElseIf Bottom > Cells.Rows.Count Or Right > Cells.Columns.Count Then
        Exit Function
    End If
    
    Set GetRangeWithPosition _
        = sh.Range(sh.Cells(Top, Left), sh.Cells(Bottom, Right))
End Function
