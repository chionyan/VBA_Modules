Attribute VB_Name = "M05_Other_Function"
'�����Z���̃A�h���X��Ԃ�
Function MargeCellAddress(ByVal Targets As Range) As String
    If Targets(1).MergeCells = True Then
        MargeCellAddress = Targets(1).MergeArea.Address
    Else
        MargeCellAddress = Targets(1).Address
    End If
End Function

'�������z�񂩂ǂ�������
Public Function IsArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

'���l��������A���t�@�x�b�g�ɕϊ��A�A���t�@�x�b�g�������琔�l�ɕϊ�����֐�
Function CNumAlp(va As Variant) As Variant
    On Error GoTo CNumAlpErr
    Dim al As String
    
    If IsNumeric(va) = True Then
        al = Cells(1, va).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        CNumAlp = Left(al, Len(al) - 1)
    Else
        CNumAlp = Range(va & "1").Column
    End If
    Exit Function
CNumAlpErr:
    CNumAlp = va
End Function

'�I��͈͂̕���������������l��Ԃ��֐�
Function CONCAT(ByVal argRng As Range) As String
    Application.Volatile
    For Each rng In argRng
        CONCAT = CONCAT + rng.Value
    Next rng
End Function

'ISFORMULA�Ɠ��l
Function ISFORMULA(ByVal argRng As Range) As Boolean
    ISFORMULA = argRng.HasFormula
End Function

'COUNT�Ɠ��l
Function wsfn_Count(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_Count = wsfn_Count + WorksheetFunction.Count(area)
    Next
    Exit Function
errProc:
    wsfn_Count = 0
End Function

'COUNTA�Ɠ��l
Function wsfn_CountA(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_CountA = wsfn_CountA + WorksheetFunction.CountA(area)
    Next
    Exit Function
errProc:
    wsfn_CountA = 0
End Function

'COUNTBLANK�Ɠ��l
Function wsfn_CountBlank(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_CountBlank = wsfn_CountBlank + WorksheetFunction.CountBlank(area)
    Next
    Exit Function
errProc:
    wsfn_CountBlank = 0
End Function

'COUNTIF�Ɠ��l
Function wsfn_CountIf(ByVal argRng As Range, ByVal args As String) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_CountIf = wsfn_CountIf + WorksheetFunction.CountIf(area, args)
    Next
    Exit Function
errProc:
    wsfn_CountIf = 0
End Function

'�v�Z���ȊO�̃J�E���g
Function wsfn_CountNotFormula(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        For Each rng In area
            If rng.Value <> "" And Not rng.HasFormula Then wsfn_CountNotFormula = wsfn_CountNotFormula + 1
        Next
    Next
    Exit Function
errProc:
    wsfn_CountNotFormula = wsfn_CountNotFormula + 1
End Function


'���p��1�o�C�g�A�S�p��2�o�C�g�Ōv�Z�����o�C�g����Ԃ�
Function LenB2(ByVal args As String) As Long
    LenB2 = LenB(StrConv(args, vbFromUnicode))
    Exit Function
End Function

'����1�Ɋ܂܂�����2�̕����̐���Ԃ�
Function StrCount(ByVal Source As String, ByVal Target As String) As Long
    Dim n As Long, cnt As Long
    Do
        n = InStr(n + 1, Source, Target)
        If n = 0 Then
            Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop
    StrCount = cnt
End Function

'�����������ł��邪���l�^�łȂ��Ƃ��A���l�^�ɕϊ����ĕԂ�
Function Cast(ByVal args)
    Cast = args
    If IsNumeric(args) Then Cast = Val(args)
End Function

'���O��`�͈͈ꗗ�擾
Sub PrintNames()
    Dim nm As Name
    For Each nm In ActiveWorkbook.Names
        Debug.Print nm.Name & ":" & nm.Value
    Next
End Sub

'�w�i�F�擾
Sub PrintBackGroundColor()
    myColor = ActiveCell.Interior.Color
    
    myR = myColor Mod 256
    myG = Int(myColor / 256) Mod 256
    myB = Int(myColor / 256 / 256)
    
    Debug.Print "�I�������Z����RGB�́A" & Chr(10) & _
            "R:" & myR & Chr(10) & _
            "G:" & myG & Chr(10) & _
            "B:" & myB & Chr(10) & _
            "�ł��B"
End Sub
