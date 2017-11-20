Attribute VB_Name = "M05_Other_Function"
'列名変換シートに入力された文字列に変換する｡
'※列名変換シートにテーブルカラム名と物理カラム名の一覧を入れておく必要あり
Function ColumnNameConversion(ByVal args As String) As String
    ColumnNameConversion = args
    For Each rng In BottomRightExtention(Sheets("列名変換").Range("A1"))
        If args = rng.Value Then
        If rng.Column = 1 Then ColumnNameConversion = rng.Offset(0, 1).Value: Exit Function
        If rng.Column = 2 Then ColumnNameConversion = rng.Offset(0, -1).Value: Exit Function
        End If
    Next rng
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

'COUNTAと同様
Function WkstCountA(ByVal argRng As Range) As Integer
    WkstCountA = WorksheetFunction.CountA(argRng)
End Function

'半角は1バイト、全角は2バイトで計算したバイト数を返す
Function LenB2(ByVal args As String) As Long
    LenB2 = LenB(StrConv(args, vbFromUnicode))
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
    If IsNumeric(args) Then Cast = val(args)
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
    
    MsgBox "選択したセルのRGBは、" & Chr(10) & _
            "R:" & myR & Chr(10) & _
            "G:" & myG & Chr(10) & _
            "B:" & myB & Chr(10) & _
            "です。"
End Sub

'dbscsetシートの入力不可セルはロックにする
Sub dbsc入力セルボタン()
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


