Attribute VB_Name = "M04_InputCheck_Function"
'���̓`�F�b�N
'******************************************************************************************
' Procedure���FinputCheck(ByVal argRng As Range) As Boolean
' �@�\�T�v�@  �F�K�{���͂̃`�F�b�N���s���B
'******************************************************************************************
Function inputCheck(ByVal argRng As Range) As Boolean
    inputCheck = True
    If Len(argRng.Value) = 0 Then
        MsgBox argRng.Value & "����͂��Ă��������B", vbCritical
        inputCheck = False
    End If
End Function
'******************************************************************************************
' Procedure���FchoiceCheck(ByVal argRng As Range, ByVal selectRng As Range) As Boolean
' �@�\�T�v�@  �F�v���_�E���I���̃`�F�b�N���s���B
'******************************************************************************************
Function choiceCheck(ByVal argRng As Range, ByVal selectRng As Range) As Boolean
    choiceCheck = True
    For Each rng In selectRng
        If argRng.Value = rng.Value Then Exit Function
    Next
    MsgBox argRng.Value & "�̓v���_�E������I��ł��������B", vbCritical
    choiceCheck = False
End Function
'******************************************************************************************
' Procedure���FbyteCheck(ByVal argRng As Range, ByVal argByte As Long) As Boolean
' �@�\�T�v�@  �F�o�C�g���̃`�F�b�N���s���B
'******************************************************************************************
Function byteCheck(ByVal argRng As Range, ByVal argByte As Long) As Boolean
    byteCheck = True
    If LenB2(argRng.Value) > argByte Then
        MsgBox argRng.Value & "��" & argByte & "�o�C�g�𒴂��Ă��܂��B(" & LenB2(argRng.Value) & ")", vbCritical
        byteCheck = False
    End If
End Function
'******************************************************************************************
' Procedure���FnumericCheck(ByVal ckRng As Range) As Boolean
' �@�\�T�v�@  �F���p�����A0�ȏ�̐����̃`�F�b�N���s���B
'******************************************************************************************
Function numericCheck(ByVal argRng As Range) As Boolean
    numericCheck = True
    If Not IsNumeric(argRng.Value) Or (LenB2(argRng.Value) <> Len(argRng.Value)) Then
        MsgBox argRng.Value & "�͔��p���l�œ��͂��Ă��������B", vbCritical
        numericCheck = False
    ElseIf argRng.Value < 0 Or Abs(argRng.Value - Int(argRng.Value)) <> 0 Then
        MsgBox argRng.Value & "��0�ȏ�̐����œ��͂��Ă��������B", vbCritical
        numericCheck = False
    End If
End Function

'���̑�
'******************************************************************************************
' Procedure���FLenB2(args As String) As Long
' �@�\�T�v�@  �F�V�X�e���̊���̃R�[�h�̕����o�C�g��Ԃ�
'******************************************************************************************
Function LenB2(ByVal args As String) As Long
    LenB2 = LenB(StrConv(args, vbFromUnicode))
End Function
'******************************************************************************************
' Procedure���FColumnNameConversion(args As String) As String
' �@�\�T�v�@  �F�񖼕ϊ��V�[�g�ɓ��͂��ꂽ������ɕϊ�����B
'               ���񖼕ϊ��V�[�g�Ƀe�[�u���J�������ƕ����J�������̈ꗗ�����Ă����K�v����
'******************************************************************************************
Function ColumnNameConversion(ByVal args As String) As String
    ColumnNameConversion = args
    Dim rng As Range
    For Each rng In BottomRightExtention(Sheets("�񖼕ϊ�").Range("A1"))
        If args = rng.Value Then
        If rng.Column = 1 Then ColumnNameConversion = rng.Offset(0, 1).Value: Exit Function
        If rng.Column = 2 Then ColumnNameConversion = rng.Offset(0, -1).Value: Exit Function
        End If
    Next rng
End Function

