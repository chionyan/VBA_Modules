Attribute VB_Name = "M04_InputCheck_Function"
'必須入力のチェックを行う｡
Function inputCheck(ByVal argRng As Range) As Boolean
    inputCheck = True
    If Len(argRng.Value) = 0 Then
        MsgBox argRng.Value & "を入力してください。", vbCritical
        inputCheck = False
    End If
End Function

'プルダウン選択のチェックを行う｡
Function choiceCheck(ByVal argRng As Range, ByVal selectRng As Range) As Boolean
    choiceCheck = True
    For Each rng In selectRng
        If argRng.Value = rng.Value Then Exit Function
    Next
    MsgBox argRng.Value & "はプルダウンから選んでください。", vbCritical
    choiceCheck = False
End Function

'バイト数のチェックを行う｡
Function byteCheck(ByVal argRng As Range, ByVal argByte As Long) As Boolean
    byteCheck = True
    If LenB2(argRng.Value) > argByte Then
        MsgBox argRng.Value & "が" & argByte & "バイトを超えています。(" & LenB2(argRng.Value) & ")", vbCritical
        byteCheck = False
    End If
End Function

'半角数字、0以上の整数のチェックを行う。
Function numericCheck(ByVal argRng As Range) As Boolean
    numericCheck = True
    If Not IsNumeric(argRng.Value) Or (LenB2(argRng.Value) <> Len(argRng.Value)) Then
        MsgBox argRng.Value & "は半角数値で入力してください。", vbCritical
        numericCheck = False
    ElseIf argRng.Value < 0 Or Abs(argRng.Value - Int(argRng.Value)) <> 0 Then
        MsgBox argRng.Value & "は0以上の整数で入力してください。", vbCritical
        numericCheck = False
    End If
End Function



