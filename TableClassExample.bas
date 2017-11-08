Attribute VB_Name = "TableClassExample"
Option Explicit
Dim tb1 As Table
Dim tb2 As Table

Sub TableClassExample()
    
    On Error Resume Next
    
    '範囲設定
    'テーブルを作成するRangeを指定してください。
    Dim rng1 As Range: Set rng1 = Range("A5:H16")
    Dim rng2 As Range: Set rng2 = Range("K14:R25")
    
     'インスタンス生成
     '先ほど指定したRangeを第2引数にしてください。
    Set tb1 = Init(New Table, rng1)
    Set tb2 = Init(New Table, rng2)
    
    ' メソッド（いろいろ）
    Debug.Print tb1.TableRange.Address
    Debug.Print tb1.TableColumn("い").Address

    '以下二つは同義
    Debug.Print tb1.TableRow("一").Address
    Debug.Print tb1.TableRow(7).Address

    '以下二つは同義
    Debug.Print tb1.TableCells("一", "ほ").Value
    Debug.Print tb1.TableCells(7, "ほ").Value

    'sampleRange2も取れる
    Debug.Print tb2.TableRange.Address
    Debug.Print tb2.TableRow("h").Address
    
End Sub

Function Init(o As Initializable, ParamArray p()) As Object
    Dim p2() As Variant
    ReDim p2(UBound(p))
    Dim i As Long
    For i = 0 To UBound(p)
        If IsObject(p(i)) Then
            Set p2(i) = p(i)
        Else
            Let p2(i) = p(i)
        End If
    Next
    Set Init = o.Init(p2)
End Function

