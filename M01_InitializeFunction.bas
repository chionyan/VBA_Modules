Attribute VB_Name = "M01_InitializeFunction"
Option Explicit
Public i As Long
Public rng As Range
Public tempRng As Range
Public dbscsetRng As Range
Public STY010_V_card As CardType
Public OPINIO_V_list As ListType

'インスタンス生成
Sub InstanceSet(Optional ByVal STY010_V_flg As Boolean = False, _
                Optional ByVal OPINIO_V_flg As Boolean = False)
    Set dbscsetRng = BottomRightExtention(Sheets("dbscset").Range("B1"))
    
    If STY010_V_flg = True Then _
    Set STY010_V_card = CardTypeInit(Range(RangeSelect(dbscsetRng, "カード型セル位置範囲", "STY010_V")))
    
    If OPINIO_V_flg = True Then _
    Set OPINIO_V_list = ListTypeInit(Range(RangeSelect(dbscsetRng, "インスタンス作成範囲", "OPINIO_V")))

End Sub





'Initializableクラスで使用するメソッドの中身
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

Function ListTypeInit(ByVal argRng As Range) As ListType
    Set ListTypeInit = Init(New ListType, argRng)
End Function

Function CardTypeInit(ByVal argRng As Range) As CardType
    Set CardTypeInit = Init(New CardType, argRng)
End Function
