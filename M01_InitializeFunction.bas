Attribute VB_Name = "M01_InitializeFunction"
Option Explicit
Public i As Long
Public rng As Range
Public Target As Range
Public tempRng As Range
Public dbscsetRng As Range
Public key As Variant

Public instDic As Object
Public WK_inst As ListType
Public dbscset_inst As ListType

'インスタンス生成
Sub InstanceSet()
    On Error Resume Next
    Dim temp_list As ListType
    Dim temp_card As CardType
    Dim sheetName As String
    Dim sheetType As String
    Dim sheetRange As Range

    Set instDic = CreateObject("scripting.dictionary")
    Set dbscsetRng = BottomRightExtention(Sheets("dbscset").Range("A1"))
    Set dbscset_inst = ListTypeInit(dbscsetRng)
    With dbscset_inst
        For Each rng In .ListCells("インスタンス生成", "dbscset見出し")
            If rng.Offset(0, 1).Value = 1 Then
                .ListHeaderColNum = 2
                sheetName = rng.Value
                sheetType = .ListCells("データ展開種類", sheetName).Value
                Set sheetRange = Range(.ListCells("インスタンス作成範囲", sheetName).Value)
                If sheetType = "リスト型" Then
                    Set temp_list = ListTypeInit(sheetRange)
                    instDic.Add sheetName, temp_list
                ElseIf sheetType = "カード型" Then
                    Set temp_card = CardTypeInit(sheetRange)
                    instDic.Add sheetName, temp_card
                End If
            End If
        Next
    End With
End Sub

'インスタンス破棄
Sub InstanceFormat()
    Set dbscset_inst = Nothing
    Set instDic = Nothing
End Sub

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

Function ListTypeInit(ByVal argRng As Range) As ListType
    Set ListTypeInit = Init(New ListType, argRng)
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
