VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Initializable

'----------メンバ変数----------
Private my_list_Sheet As Worksheet
Private my_list_Range As Range
Private my_list_HeaderRowNum As Integer
Private my_list_HeaderColNum As Integer
Private my_list_HeaderRow As Range
Private my_list_HeaderCol As Range
Private my_list_cfg As Object

'----------コンストラクタ----------
Private Function Initializable_Init(p() As Variant) As Object
    On Error Resume Next
    
    Set my_list_Range = p(0)
    If UBound(p) - LBound(p) + 1 = 2 Then
        my_list_HeaderRowNum = my_list_Range(1).Row + p(1) - 1
    Else
        my_list_HeaderRowNum = my_list_Range(1).Row
    End If
    If UBound(p) - LBound(p) + 1 = 3 Then
        my_list_HeaderColNum = my_list_Range(1).Column + p(2) - 1
    Else
        my_list_HeaderColNum = my_list_Range(1).Column
    End If
    Set my_list_Sheet = my_list_Range.Parent
    Set my_list_HeaderRow = RowSelect(my_list_Range, my_list_HeaderRowNum)
    Set my_list_HeaderCol = ColSelect(my_list_Range, my_list_HeaderColNum)

    'コンフィグ
    Set my_list_cfg = CreateObject("Scripting.Dictionary")
    For Each rng In ColSelect(dbscsetRng, "dbscset見出し")
        my_list_cfg.Add rng.Value, RangeSelect(dbscsetRng, rng.Value, my_list_Sheet.Name).Value
    Next
    
    Set Initializable_Init = Me
End Function

'----------Getter----------
Public Property Get ListSheet() As Worksheet
    Set ListSheet = my_list_Sheet
End Property

Public Property Get ListRange() As Range
    Set ListRange = my_list_Range
End Property

Public Property Get ListHeaderRow() As Range
    Set ListHeaderRow = my_list_HeaderRow
End Property

Public Property Get ListHeaderCol() As Range
    Set ListHeaderCol = my_list_HeaderCol
End Property

'----------関数----------
Public Function ListRow(ByVal rowObject) As Range
    Set ListRow = RowSelect(my_list_Range, rowObject, my_list_HeaderColNum)
End Function

Public Function ListColumn(ByVal colObject) As Range
    Set ListColumn = ColSelect(my_list_Range, colObject, my_list_HeaderRowNum)
End Function

Public Function ListCells(ByVal rowObject, ByVal colObject) As Range
    Set ListCells = RangeSelect(my_list_Range, rowObject, colObject, my_list_HeaderColNum, my_list_HeaderRowNum)
End Function

Public Function ListRowNum(ByVal rowObject) As Long
    ListRowNum = Me.ListRow(rowObject).Row
End Function

Public Function ListColumnNum(ByVal colObject) As Long
    ListColumnNum = Me.ListRow(colObject).Column
End Function

Public Function ListMinColNum() As Long
    ListMinColNum = RangeMinColNum(my_list_Range)
End Function

Public Function ListMaxColNum() As Long
    ListMaxColNum = RangeMaxColNum(my_list_Range)
End Function

Public Function ListMinRowNum() As Long
    ListMinRowNum = RangeMinRowNum(my_list_Range)
End Function

Public Function ListMaxRowNum() As Long
    ListMaxRowNum = RangeMaxRowNum(my_list_Range)
End Function

Public Function ListConfig(ByVal args As String) As String
    ListConfig = my_list_cfg(args)
End Function

Public Function ListConfigs() As Object
    Set ListConfigs = my_list_cfg
End Function

'----------サブルーチン----------
'行の削除
Public Sub ListRowsDelete(ByVal strRow As Long)
    With Me.ListSheet
        Range(.Rows(strRow), .Rows(Rows.Count)).Delete
    End With
End Sub

'範囲のコピぺ
Public Sub ListRowsCopyAndPaste(ByVal cpSrcRng As Range, ByVal strRow As Integer, ByVal endRow As Integer)
    With Me.ListSheet
        cpSrcRng.Copy
        Dim rowNum As Integer
        For rowNum = strRow To endRow
            RowSelect(cpSrcRng, rowNum).PasteSpecial Paste:=xlPasteFormats
            RowSelect(cpSrcRng, rowNum).PasteSpecial Paste:=xlPasteValidation
            For Each rng In cpSrcRng
                If ISFORMULA(rng) Then RowSelect(rng, rowNum).FormulaR1C1 = rng.FormulaR1C1
            Next
        Next
    End With
    Application.CutCopyMode = False
End Sub

'オートフィルタ
Public Sub ListAutoFilter(Optional ByVal offsetRow As Integer = 0)
    Me.ListRange.Offset(offsetRow).Resize(Me.ListRange.Rows.Count + offsetRow).AutoFilter
End Sub

