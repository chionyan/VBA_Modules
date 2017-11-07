Attribute VB_Name = "M04_InputCheck_Function"
'入力チェック
'******************************************************************************************
' Procedure名：inputCheck(ByVal argRng As Range) As Boolean
' 機能概要　  ：必須入力のチェックを行う。
'******************************************************************************************
Function inputCheck(ByVal argRng As Range) As Boolean
    inputCheck = True
    If Len(argRng.Value) = 0 Then
        MsgBox argRng.Value & "を入力してください。", vbCritical
        inputCheck = False
    End If
End Function
'******************************************************************************************
' Procedure名：choiceCheck(ByVal argRng As Range, ByVal selectRng As Range) As Boolean
' 機能概要　  ：プルダウン選択のチェックを行う。
'******************************************************************************************
Function choiceCheck(ByVal argRng As Range, ByVal selectRng As Range) As Boolean
    choiceCheck = True
    For Each rng In selectRng
        If argRng.Value = rng.Value Then Exit Function
    Next
    MsgBox argRng.Value & "はプルダウンから選んでください。", vbCritical
    choiceCheck = False
End Function
'******************************************************************************************
' Procedure名：byteCheck(ByVal argRng As Range, ByVal argByte As Long) As Boolean
' 機能概要　  ：バイト数のチェックを行う。
'******************************************************************************************
Function byteCheck(ByVal argRng As Range, ByVal argByte As Long) As Boolean
    byteCheck = True
    If LenB2(argRng.Value) > argByte Then
        MsgBox argRng.Value & "が" & argByte & "バイトを超えています。(" & LenB2(argRng.Value) & ")", vbCritical
        byteCheck = False
    End If
End Function
'******************************************************************************************
' Procedure名：numericCheck(ByVal ckRng As Range) As Boolean
' 機能概要　  ：半角数字、0以上の整数のチェックを行う。
'******************************************************************************************
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

'その他
'******************************************************************************************
' Procedure名：LenB2(args As String) As Long
' 機能概要　  ：システムの既定のコードの文字バイトを返す
'******************************************************************************************
Function LenB2(ByVal args As String) As Long
    LenB2 = LenB(StrConv(args, vbFromUnicode))
End Function
'******************************************************************************************
' Procedure名：ColumnNameConversion(args As String) As String
' 機能概要　  ：列名変換シートに入力された文字列に変換する。
'               ※列名変換シートにテーブルカラム名と物理カラム名の一覧を入れておく必要あり
'******************************************************************************************
Function ColumnNameConversion(ByVal args As String) As String
    ColumnNameConversion = args
    Dim rng As Range
    For Each rng In BottomRightExtention(Sheets("列名変換").Range("A1"))
        If args = rng.Value Then
        If rng.Column = 1 Then ColumnNameConversion = rng.Offset(0, 1).Value: Exit Function
        If rng.Column = 2 Then ColumnNameConversion = rng.Offset(0, -1).Value: Exit Function
        End If
    Next rng
End Function

