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

'----------メンバ変数----------
Private my_card_cellsList As ListType
Private my_card_Sheet As Worksheet
Private my_card_Range As Range
Private my_card_NameCells As Object
Private my_card_cfg As Object

'----------コンストラクタ----------
Private Function Initializable_Init(p() As Variant) As Object
    On Error Resume Next
    
    '個別セルセット
    Set my_card_cellsList = ListTypeInit(p(0))
    With my_card_cellsList
        For i = .ListMinRowNum + 1 To .ListMaxRowNum
            Set my_card_Range = Union2(my_card_Range, Range(.ListCells(i, "範囲").Value))
        Next
    End With
    
    Set my_card_Sheet = my_card_Range.Parent

    'コンフィグ
    Set my_card_cfg = CreateObject("Scripting.Dictionary")
    For Each rng In ColSelect(dbscsetRng, "dbscset見出し")
        my_card_cfg.Add rng.Value, RangeSelect(dbscsetRng, rng.Value, my_card_Sheet.Name).Value
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

'----------関数----------
Public Function CardConfig(ByVal args As String) As String
    CardConfig = my_card_cfg(args)
End Function

Public Function CardConfigs() As Object
    Set CardConfigs = my_card_cfg
End Function

Public Function CardNameCell(ByVal args As String) As Range
    Set CardNameCell = Range(my_card_cellsList.ListCells(args, "範囲"))
End Function

Public Function CardNameStyle(ByVal args As String) As String
    CardNameStyle = my_card_cellsList.ListCells(args, "セルのスタイル").Value
End Function

'----------サブルーチン----------
'値のクリア
Public Sub CardClearContents()
    Me.CardRange.ClearContents
End Sub

'セルのスタイルセット
Public Sub CardStyleSetting()
    With my_card_cellsList
        For Each rng In .ListHeaderCol.Offset(1).Resize(.ListHeaderCol.Rows.Count - 1)
            Dim cellName As String: cellName = rng.Value
            Me.CardNameCell(cellName).Style = Me.CardNameStyle(cellName)
        Next
    End With
End Sub

