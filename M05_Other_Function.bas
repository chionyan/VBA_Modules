Attribute VB_Name = "M05_Other_Function"
'結合セルのアドレスを返す
Function MargeCellAddress(ByVal Targets As Range) As String
    If Targets(1).MergeCells = True Then
        MargeCellAddress = Targets(1).MergeArea.Address
    Else
        MargeCellAddress = Targets(1).Address
    End If
End Function

'引数が配列かどうか判別
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

'数値だったらアルファベットに変換、アルファベットだったら数値に変換する関数
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

'選択範囲の文字列を結合した値を返す関数
Function CONCAT(ByVal argRng As Range) As String
    Application.Volatile
    For Each rng In argRng
        CONCAT = CONCAT + rng.Value
    Next rng
End Function

'ISFORMULAと同様
Function ISFORMULA(ByVal argRng As Range) As Boolean
    ISFORMULA = argRng.HasFormula
End Function

'COUNTと同様
Function wsfn_Count(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_Count = wsfn_Count + WorksheetFunction.Count(area)
    Next
    Exit Function
errProc:
    wsfn_Count = 0
End Function

'COUNTAと同様
Function wsfn_CountA(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_CountA = wsfn_CountA + WorksheetFunction.CountA(area)
    Next
    Exit Function
errProc:
    wsfn_CountA = 0
End Function

'COUNTBLANKと同様
Function wsfn_CountBlank(ByVal argRng As Range) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_CountBlank = wsfn_CountBlank + WorksheetFunction.CountBlank(area)
    Next
    Exit Function
errProc:
    wsfn_CountBlank = 0
End Function

'COUNTIFと同様
Function wsfn_CountIf(ByVal argRng As Range, ByVal args As String) As Long
    On Error GoTo errProc
    For Each area In argRng
        wsfn_CountIf = wsfn_CountIf + WorksheetFunction.CountIf(area, args)
    Next
    Exit Function
errProc:
    wsfn_CountIf = 0
End Function

'計算式以外のカウント
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


'半角は1バイト、全角は2バイトで計算したバイト数を返す
Function LenB2(ByVal args As String) As Long
    LenB2 = LenB(StrConv(args, vbFromUnicode))
    Exit Function
End Function

'引数1に含まれる引数2の文字の数を返す
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

'引数が数字であるが数値型でないとき、数値型に変換して返す
Function Cast(ByVal args)
    Cast = args
    If IsNumeric(args) Then Cast = Val(args)
End Function

'名前定義範囲一覧取得
Sub PrintNames()
    Dim nm As Name
    For Each nm In ActiveWorkbook.Names
        Debug.Print nm.Name & ":" & nm.Value
    Next
End Sub

'背景色取得
Sub PrintBackGroundColor()
    myColor = ActiveCell.Interior.Color
    
    myR = myColor Mod 256
    myG = Int(myColor / 256) Mod 256
    myB = Int(myColor / 256 / 256)
    
    Debug.Print "選択したセルのRGBは、" & Chr(10) & _
            "R:" & myR & Chr(10) & _
            "G:" & myG & Chr(10) & _
            "B:" & myB & Chr(10) & _
            "です。"
End Sub
