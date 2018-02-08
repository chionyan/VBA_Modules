Attribute VB_Name = "M01_InitializeFunction"
Option Explicit
Public i As Long
Public rng As Range
Public Target As Range
Public key As Variant

Function ToDataType(ByVal argInst As Variant) As DataType
    If TypeName(argInst) = "ListType" Or TypeName(argInst) = "CardType" Then
        Set ToDataType = argInst
    End If
End Function

Function ToListType(ByVal argInst As DataType) As ListType
    Set ToListType = argInst
End Function

Function ToCardType(ByVal argInst As DataType) As CardType
    Set ToCardType = argInst
End Function

Function ListTypeInit(ByVal argRng As Range, Optional headerRowNum = 0, Optional headerColNum = 0) As ListType
    If headerRowNum = 0 Then headerRowNum = argRng(1).Row
    If headerColNum = 0 Then headerColNum = argRng(1).Column
    Set ListTypeInit = Init(New ListType, argRng, headerRowNum, headerColNum)
End Function

Function CardTypeInit(ByVal argRng As Range) As CardType
    Set CardTypeInit = Init(New CardType, argRng)
End Function

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
