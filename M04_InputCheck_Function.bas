Attribute VB_Name = "M04_InputCheck_Function"
'�K�{���͂̃`�F�b�N���s���
Function inputCheck(ByVal argRng As Range) As Boolean
    inputCheck = True
    If Len(argRng.Value) = 0 Then
        MsgBox argRng.Value & "����͂��Ă��������B", vbCritical
        inputCheck = False
    End If
End Function

'�v���_�E���I���̃`�F�b�N���s���
Function choiceCheck(ByVal argRng As Range, ByVal selectRng As Range) As Boolean
    choiceCheck = True
    For Each rng In selectRng
        If argRng.Value = rng.Value Then Exit Function
    Next
    MsgBox argRng.Value & "�̓v���_�E������I��ł��������B", vbCritical
    choiceCheck = False
End Function

'�o�C�g���̃`�F�b�N���s���
Function byteCheck(ByVal argRng As Range, ByVal argByte As Long) As Boolean
    byteCheck = True
    If LenB2(argRng.Value) > argByte Then
        MsgBox argRng.Value & "��" & argByte & "�o�C�g�𒴂��Ă��܂��B(" & LenB2(argRng.Value) & ")", vbCritical
        byteCheck = False
    End If
End Function

'���p�����A0�ȏ�̐����̃`�F�b�N���s���B
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



