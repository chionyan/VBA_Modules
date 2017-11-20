Attribute VB_Name = "M03_RangeOperation_Function"
Dim resultRng As Range

Function RangeLTopCell(ByVal argRng As Range) As Range
    Set RangeLTopCell = argRng(1).Offset(0, 0)
    For Each Area In argRng.Areas
        Set tempRng = Area(1).Offset(0, 0)
        If RangeLTopCell.Column > tempRng.Column Then Set RangeLTopCell = tempRng
    Next
End Function

Function RangeRTopCell(ByVal argRng As Range) As Range
    Set RangeRTopCell = argRng(1).Offset(0, argRng.Columns.Count - 1)
    For Each Area In argRng.Areas
        Set tempRng = Area(1).Offset(0, Area.Columns.Count - 1)
        If RangeRTopCell.Column < tempRng.Column Then Set RangeRTopCell = tempRng
    Next
End Function

Function RangeLBtmCell(ByVal argRng As Range) As Range
    Set RangeLBtmCell = argRng(1).Offset(argRng.Rows.Count - 1, 0)
    For Each Area In argRng.Areas
        Set tempRng = Area(1).Offset(Area.Rows.Count - 1, 0)
        If RangeLBtmCell.Column > tempRng.Column Then Set RangeLBtmCell = tempRng
    Next
End Function

Function RangeRBtmCell(ByVal argRng As Range) As Range
    Set RangeRBtmCell = argRng(1).Offset(argRng.Rows.Count - 1, argRng.Columns.Count - 1)
    For Each Area In argRng.Areas
        Set tempRng = Area(1).Offset(Area.Rows.Count - 1, Area.Columns.Count - 1)
        If RangeRBtmCell.Column < tempRng.Column Then Set RangeRBtmCell = tempRng
    Next
End Function

Function RangeMinColNum(ByVal argRng As Range) As Long
    RangeMinColNum = RangeLTopCell(argRng).Column
End Function

Function RangeMaxColNum(ByVal argRng As Range) As Long
    RangeMaxColNum = RangeRBtmCell(argRng).Column
End Function

Function RangeMinRowNum(ByVal argRng As Range) As Long
    RangeMinRowNum = RangeLTopCell(argRng).Row
End Function

Function RangeMaxRowNum(ByVal argRng As Range) As Long
    RangeMaxRowNum = RangeRBtmCell(argRng).Row
End Function

Function RowSelect(ByVal argRange As Range, ByVal rowObj, Optional ByVal headerColNum = 0) As Range
    '����1�F�Ώ۔͈́A'����2:�����p�^�[���A'����3:�����Ώۂ̗�w��
    '�y�����p�^�[���z
    '�E���p�p�����ƕ����^�i���K�\���j�̕��p�\�B�i���p�p�����������ΏۂƂ���ꍇ�͐��K�\����"^1$"�ȂǂƎw��B�j
    '�E�u����:�����v�Ł����`�����̍s���w��B�u����,�����v�Ł����Ɓ����̍s���w��i,�͕����g�p�\�j
    '�y�����Ώۂ̗�w��z
    '�E�ȗ��̏ꍇ�Ώ۔͈͂̈�ԍ��̗���w��B
    '�E�����^�Ŏw�肷��ƑΏ۔͈͓��ł̑��Η񐔂ɂȂ�܂��B

    If headerColNum = 0 Then
        Set headerCol = argRange.Resize(, 1)
    Else
        If TypeName(headerColNum) = "String" Then
            Set headerCol = argRange.Offset(, headerColNum - 1).Resize(, 1)
        Else
            Set headerCol = argRange.Offset(, headerColNum - argRange(1).Column).Resize(, 1)
        End If
    End If
    
    For Each ptn In Split(rowObj, ",")
        If StrCount(ptn, ":") = 0 Then s_ptn = Cast(ptn): e_ptn = Cast(ptn)
        If StrCount(ptn, ":") = 1 Then s_ptn = Cast(Split(ptn, ":")(0)): e_ptn = Cast(Split(ptn, ":")(1))
        
        For Each temp_ptn In Array(s_ptn, e_ptn)
            If Not IsNumeric(temp_ptn) Then
                Set temp_ptnRng = ColResize(FilterRng(headerCol, temp_ptn), argRange(1).Column, argRange(argRange.Count).Column)
            Else
                Set temp_ptnRng = RowResize(argRange, temp_ptn, temp_ptn)
            End If
            If temp_ptn = s_ptn Then Set s_ptnRng = temp_ptnRng
            If temp_ptn = e_ptn Then Set e_ptnRng = temp_ptnRng
        Next
        
        If StrCount(ptn, ":") = 0 Then Set RowSelect = Union2(RowSelect, Union2(s_ptnRng, e_ptnRng))
        If StrCount(ptn, ":") = 1 Then Set RowSelect = Union2(RowSelect, Range(s_ptnRng, e_ptnRng))
    Next

End Function

Function ColSelect(ByVal argRange As Range, ByVal colObj, Optional ByVal headerRowNum = 0) As Range
    '����1�F�Ώ۔͈́A'����2:�����p�^�[���A'����3:�����Ώۂ̍s�w��
    '�ڂ������e��RowSelect�Ɠ��l�B
    '����2�̓A���t�@�x�b�g�ɂ��񖼂̎w����\�B(ex."J"�Ȃ�)

    Dim alpha As Object: Set alpha = CreateObject("VBScript.RegExp")
    alpha.Pattern = "^[A-Z]{1,2}$"

    If headerRowNum = 0 Then
        Set headerRow = argRange.Resize(1)
    Else
        If TypeName(headerRowNum) = "String" Then
            Set headerRow = argRange.Offset(headerRowNum - 1).Resize(1)
        Else
            Set headerRow = argRange.Offset(headerRowNum - argRange(1).Row).Resize(1)
        End If
    End If
    
    For Each ptn In Split(colObj, ",")
        If StrCount(ptn, ":") = 0 Then s_ptn = Cast(ptn): e_ptn = Cast(ptn)
        If StrCount(ptn, ":") = 1 Then s_ptn = Cast(Split(ptn, ":")(0)): e_ptn = Cast(Split(ptn, ":")(1))

        For Each temp_ptn In Array(s_ptn, e_ptn)
            If Not IsNumeric(temp_ptn) And Not alpha.test(temp_ptn) Then
                Set temp_ptnRng = RowResize(FilterRng(headerRow, temp_ptn), argRange(1).Row, argRange(argRange.Count).Row)
            Else
                If Not IsNumeric(temp_ptn) Then temp_ptn = CNumAlp(temp_ptn)
                Set temp_ptnRng = ColResize(argRange, temp_ptn, temp_ptn)
            End If
            If temp_ptn = s_ptn Or temp_ptn = CNumAlp(s_ptn) Then Set s_ptnRng = temp_ptnRng
            If temp_ptn = e_ptn Or temp_ptn = CNumAlp(e_ptn) Then Set e_ptnRng = temp_ptnRng
        Next
        
        If StrCount(ptn, ":") = 0 Then Set ColSelect = Union2(ColSelect, Union2(s_ptnRng, e_ptnRng))
        If StrCount(ptn, ":") = 1 Then Set ColSelect = Union2(ColSelect, Range(s_ptnRng, e_ptnRng))
    Next
    
End Function

'RowSelect��ColSelect��g�ݍ��킹�����́B
Function RangeSelect(ByVal argRange As Range, ByVal rowObj, ByVal colObj, Optional ByVal headerColNum = 0, Optional ByVal headerRowNum = 0) As Range
    Set RangeSelect = Intersect2(RowSelect(argRange, rowObj, headerColNum), ColSelect(argRange, colObj, headerRowNum))
End Function

'�i����1�̏ꍇ�j����1�͈̔͂�����2�̍s�T�C�Y�Ƀ��T�C�Y����B
'�i����2�̏ꍇ�j����1�͈̔͂�����2�������3�̍s���Ƀ��T�C�Y����B
Function RowResize(ByVal argRng As Range, ByVal argi1 As Integer, Optional ByVal argi2 As Integer) As Range
    For Each Area In argRng.Areas
        Set resultRng = Nothing
        If argi2 = 0 Then
            Set resultRng = Resize2(Area, argi1)
        Else
            Set resultRng = Resize2(Area.Offset(argi1 - argRng(1).Row), argi2 - argi1 + 1)
        End If
        Set RowResize = Union2(RowResize, resultRng)
    Next
End Function

'�i����1�̏ꍇ�j����1�͈̔͂�����2�̗�T�C�Y�Ƀ��T�C�Y����B
'�i����2�̏ꍇ�j����1�͈̔͂�����2�������3�̗񐔂Ƀ��T�C�Y����B
Function ColResize(ByVal argRng As Range, ByVal argi1 As Integer, Optional ByVal argi2 As Integer) As Range
    For Each Area In argRng.Areas
        Set resultRng = Nothing
        If argi2 = 0 Then
            Set resultRng = Resize2(Area, , argi1)
        Else
            Set resultRng = Resize2(Area.Offset(, argi1 - argRng(1).Column), , argi2 - argi1 + 1)
        End If
        Set ColResize = Union2(ColResize, resultRng)
    Next
End Function

'�����͈͂̃f�[�^�̂����[�܂Ŕ͈͂��g�������͈͂�Ԃ�
Function TopExtention(ByVal argRng As Range) As Range
    Dim minRowNum As Long
    For Each Area In argRng.Areas
        minRowNum = Rows.Count
        Set resultRng = Nothing
        For Each rng In Area.Resize(1)
            If rng.Parent.Cells(1, rng.Column).End(xlDown).Row < minRowNum Then
                minRowNum = rng.Parent.Cells(1, rng.Column).End(xlDown).Row
                Set resultRng = Range(Area, rng.Parent.Cells(minRowNum, rng.Column))
            End If
        Next
        Set TopExtention = Union2(TopExtention, resultRng)
    Next
End Function

'�����͈͂̃f�[�^�̂��鉺�[�܂Ŕ͈͂��g�������͈͂�Ԃ�
Function BottomExtention(ByVal argRng As Range) As Range
    Dim maxRowNum As Long
    For Each Area In argRng.Areas
        maxRowNum = 0
        Set resultRng = Nothing
        For Each rng In Area.Resize(1)
            If rng.Parent.Cells(Rows.Count, rng.Column).End(xlUp).Row > maxRowNum Then
                maxRowNum = rng.Parent.Cells(Rows.Count, rng.Column).End(xlUp).Row
                Set resultRng = Range(Area, rng.Parent.Cells(maxRowNum, rng.Column))
            End If
        Next
        Set BottomExtention = Union2(BottomExtention, resultRng)
    Next
End Function

'�����͈͂̃f�[�^�̂��鍶�[�܂Ŕ͈͂��g�������͈͂�Ԃ�
Function LeftExtention(ByVal argRng As Range) As Range
    Dim minColNum As Long
    For Each Area In argRng.Areas
        minColNum = Columns.Count
        Set resultRng = Nothing
        For Each rng In Area.Resize(, 1)
            If rng.Parent.Cells(rng.Row, 1).End(xlToRight).Column < minColNum Then
                minColNum = rng.Parent.Cells(rng.Row, 1).End(xlToRight).Column
                Set resultRng = Range(Area, rng.Parent.Cells(rng.Row, minColNum))
            End If
        Next
        Set LeftExtention = Union2(LeftExtention, resultRng)
    Next
End Function

'�����͈͂̃f�[�^�̂���E�[�܂Ŕ͈͂��g�������͈͂�Ԃ�
Function RightExtention(ByVal argRng As Range) As Range
    Dim maxColNum As Long
    For Each Area In argRng.Areas
        maxColNum = 0
        Set resultRng = Nothing
        For Each rng In Area.Resize(, 1)
            If rng.Parent.Cells(rng.Row, Columns.Count).End(xlToLeft).Column > maxColNum Then
                maxColNum = rng.Parent.Cells(rng.Row, Columns.Count).End(xlToLeft).Column
                Set resultRng = Range(Area, rng.Parent.Cells(rng.Row, maxColNum))
            End If
        Next
        Set RightExtention = Union2(RightExtention, resultRng)
    Next
End Function

'�����͈͂̃f�[�^�̂���E���[�܂Ŕ͈͂��g�������͈͂�Ԃ�
Function BottomRightExtention(ByVal argRng As Range) As Range
    Set BottomRightExtention = BottomExtention(RightExtention(argRng))
End Function

'����1�͈͓̔��ň���2�̐��K�\���Ɉ�v����͈݂͂̂��t�B���^�����O���Ĕ͈͂�Ԃ��֐�
Function FilterRng(ByVal argRng As Range, ByVal args As String) As Range
    Set FilterRng = Nothing
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = args
    For Each rng In argRng
        If Not IsError(rng.Value) Then
            If re.test(rng.Value) Then
                Set FilterRng = Union2(FilterRng, rng)
            End If
        End If
    Next
End Function

' ����1�͈͓̔��ŁA�ҏW�\�ȃZ����Ԃ��֐�
Function InpuListRng(ByVal argRng As Range) As Range
    Set InpuListRng = Nothing
    For Each rng In argRng
        If rng.Locked = False Then
            Set InpuListRng = Union2(InpuListRng, rng)
        End If
    Next rng
End Function

'�����̈�ɑΉ�����Resize
Function Resize2(ByVal argRng As Range, Optional ByVal rowSize As Long = 0, Optional ByVal colSize As Long = 0) As Range
    For Each Area In argRng.Areas
        If rowSize = 0 Then Set Resize2 = Union2(Resize2, Area.Resize(, colSize))
        If colSize = 0 Then Set Resize2 = Union2(Resize2, Area.Resize(rowSize))
    Next
End Function

' �����̃Z�� ArgList �̘a�W����Ԃ�
Function Union2(ParamArray argList() As Variant) As Range
    Dim buf As Range
    Dim i As Long
    For i = 0 To UBound(argList)
        If TypeName(argList(i)) = "Range" Then
            If buf Is Nothing Then
                Set buf = argList(i)
            Else
                Set buf = Application.Union(buf, argList(i))
            End If
        End If
    Next
    Set Union2 = buf
End Function

'�����̃Z�� ArgList �̐ϏW����Ԃ�
Function Intersect2(ParamArray argList() As Variant) As Range
    Dim buf As Range
    Dim i As Long
    For i = 0 To UBound(argList)
        If Not TypeName(argList(i)) = "Range" Then
            Exit Function
        ElseIf buf Is Nothing Then
            Set buf = argList(i)
        Else
            Set buf = Application.Intersect(buf, argList(i))
        End If
        If buf Is Nothing Then Exit Function
    Next
    Set Intersect2 = buf
End Function

' SourceRange ���� ArgList ���������������W����Ԃ�
' (SourceRange �� ���]���� ArgList �Ƃ̐ϏW����Ԃ�)
Function Except2 _
    (ByRef SourceRange As Variant, ParamArray argList() As Variant) As Range
    If TypeName(SourceRange) = "Range" Then
        Dim buf As Range
        Set buf = SourceRange
        Dim i As Long
        For i = 0 To UBound(argList)
            If TypeName(argList(i)) = "Range" Then
                Set buf = Intersect2(buf, Invert2(argList(i)))
            End If
        Next
        Set Except2 = buf
    End If
End Function

'SourceRange �̑I��͈͂𔽓]����
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
        
        Dim RangeLeft   As Range
        Set RangeLeft = GetRangeWithPosition(sh, _
            sh.Cells.Row, sh.Cells.Column, sh.Rows.Count, AreaLeft - 1)

        Dim RangeRight  As Range
        Set RangeRight = GetRangeWithPosition(sh, _
            sh.Cells.Row, AreaRight + 1, sh.Rows.Count, sh.Columns.Count)
        
        Dim RangeTop    As Range
        Set RangeTop = GetRangeWithPosition(sh, _
            sh.Cells.Row, AreaLeft, AreaTop - 1, AreaRight)
        
        Dim RangeBottom As Range
        Set RangeBottom = GetRangeWithPosition(sh, _
            AreaBottom + 1, AreaLeft, sh.Rows.Count, AreaRight)
        
        Set buf = Intersect2(buf, _
            Union2(RangeLeft, RangeRight, RangeTop, RangeBottom))
    Next
    Set Invert2 = buf
End Function

'�l�����w�肵�� Range �𓾂�
Function GetRangeWithPosition(ByRef sh As Worksheet, _
    ByVal Top As Long, ByVal Left As Long, _
    ByVal Bottom As Long, ByVal Right As Long) As Range
    
    '��������
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
