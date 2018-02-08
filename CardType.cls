VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CardType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Initializable
Implements DataType

'----------メンバ変数----------
Private my_card_CellsList As ListType
Private my_card_Sheet As Worksheet
Private my_card_Range As Range
Private my_card_NameCells As Object
Private my_card_cfg As Object

'----------コンストラクタ----------
Private Function Initializable_Init(p() As Variant) As Object
    On Error Resume Next
    
    Set my_card_CellsList = ListTypeInit(p(0))
    With my_card_CellsList
        For i = .ListMinRowNum + 1 To .ListMaxRowNum
            Set my_card_Range = Union2(my_card_Range, Range(.ListCells(i, "範囲").Value))
        Next
    End With
    Set my_card_Sheet = my_card_Range.Parent

    Set my_card_cfg = CreateObject("Scripting.Dictionary")
    For Each rng In RangeSelect(dbscsetRng, "^行$:^その他$", "dbscset見出し")
        my_card_cfg.Add rng.Value, RangeSelect(dbscsetRng, rng.Value, my_card_Sheet.Name, 2).Value
    Next
    
    Set Initializable_Init = Me
End Function


'----------Getter----------
Public Property Get CardSheet() As Worksheet
    Set CardSheet = my_card_Sheet
End Property

Public Property Get CardRange() As Range
    Set CardRange = my_card_Range
End Property

Public Property Get CardCellsList() As ListType
    Set CardCellsList = my_card_CellsList
End Property

'@Override
Public Property Get DataType_DataSheet() As Worksheet
    Set DataType_DataSheet = Me.CardSheet
End Property

Public Property Get DataType_DataRange() As Range
    Set DataType_DataRange = Me.CardRange
End Property


'----------Setter----------
Public Property Let CardSheet(ByVal new_Sheet As Worksheet)
    Set my_card_Sheet = new_Sheet
End Property

Public Property Let CardRange(ByVal new_Range As Range)
    Set my_card_Range = new_Range
End Property

Public Property Let CardCellsList(ByVal new_list As ListType)
    Set my_card_CellsList = new_list
End Property

'@Override
Public Property Let DataType_DataSheet(ByVal new_Sheet As Worksheet)
    Set DataType_DataSheet = new_Sheet
End Property

Public Property Let DataType_DataRange(ByVal new_Range As Range)
    Set DataType_DataRange = new_Range
End Property


'----------関数----------
Public Function CardConfig(ByVal args As String) As String
    CardConfig = my_card_cfg(args)
End Function

Public Function CardConfigs() As Object
    Set CardConfigs = my_card_cfg
End Function

Public Function CardName(ByVal argRng As Range) As String
    Dim argRngAddress As String
    argRngAddress = Me.CardSheet.Name & "!" & Replace(argRng.Address, "$", "")
    With my_card_CellsList
        CardName = Intersect2(.ListColumn("名前"), .ListRow(argRngAddress, "範囲"))(1).Value
    End With
End Function

Public Function CardNameCell(ByVal args As String) As Range
    Set CardNameCell = Range(my_card_CellsList.ListCells(args, "範囲"))
End Function

Public Function CardNameStyle(ByVal args As String) As String
    CardNameStyle = my_card_CellsList.ListCells(args, "セルのスタイル").Value
End Function

'@Override
Public Function DataType_DataConfig(ByVal args As String) As String
    DataType_DataConfig = Me.CardConfig(args)
End Function

Public Function DataType_DataConfigs() As Object
    Set DataType_DataConfigs = Me.CardConfigs
End Function
