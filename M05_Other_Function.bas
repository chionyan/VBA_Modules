Attribute VB_Name = "M05_Other_Function"
'�񖼕ϊ��V�[�g�ɓ��͂��ꂽ������ɕϊ�����
'���񖼕ϊ��V�[�g�Ƀe�[�u���J�������ƕ����J�������̈ꗗ�����Ă����K�v����
Function ColumnNameConversion(ByVal args As String) As String
    ColumnNameConversion = args
    For Each rng In BottomRightExtention(Sheets("�񖼕ϊ�").Range("A1"))
        If args = rng.Value Then
        If rng.Column = 1 Then ColumnNameConversion = rng.Offset(0, 1).Value: Exit Function
        If rng.Column = 2 Then ColumnNameConversion = rng.Offset(0, -1).Value: Exit Function
        End If
    Next rng
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

'COUNTA�Ɠ��l
Function WkstCountA(ByVal argRng As Range) As Integer
    WkstCountA = WorksheetFunction.CountA(argRng)
End Function

'���p��1�o�C�g�A�S�p��2�o�C�g�Ōv�Z�����o�C�g����Ԃ�
Function LenB2(ByVal args As String) As Long
    LenB2 = LenB(StrConv(args, vbFromUnicode))
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
    If IsNumeric(args) Then Cast = val(args)
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
    
    MsgBox "�I�������Z����RGB�́A" & Chr(10) & _
            "R:" & myR & Chr(10) & _
            "G:" & myG & Chr(10) & _
            "B:" & myB & Chr(10) & _
            "�ł��B"
End Sub

'dbscset�V�[�g�̓��͕s�Z���̓��b�N�ɂ���
Sub dbsc���̓Z���{�^��()
    Dim inputRange As Range
    Set inputRange = BottomRightExtention(Sheets("dbscset").Range("C2"))
    
    Call BeforeDataSet
    For Each rng In inputRange
        If rng.HasFormula Then
            rng.Locked = True
            rng.Interior.Color = RGB(191, 191, 191)
        Else
            rng.Locked = False
            rng.Interior.Color = RGB(255, 255, 255)
        End If
    Next
    Call AfterDataSet
End Sub


