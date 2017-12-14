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
Implements DataType

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
    Set my_list_Sheet = my_list_Range.Parent
    my_list_HeaderRowNum = my_list_Range(1).Row
    my_list_HeaderColNum = my_list_Range(1).Column
    
    If UBound(p) - LBound(p) + 1 = 2 Then my_list_HeaderRowNum = my_list_Range(1).Row + p(1) - 1
    If UBound(p) - LBound(p) + 1 = 3 Then my_list_HeaderColNum = my_list_Range(1).Column + p(2) - 1
    
    Set my_list_cfg = CreateObject("Scripting.Dictionary")
    For Each rng In RangeSelect(dbscsetRng, "^行$:^その他$", "dbscset見出し")
        my_list_cfg.Add rng.Value, RangeSelect(dbscsetRng, rng.Value, my_list_Sheet.Name, 2).Value
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

Public Property Get ListHeaderRowNum() As Long
    ListHeaderRowNum = my_list_HeaderRowNum
End Property

Public Property Get ListHeaderColNum() As Long
    ListHeaderColNum = my_list_HeaderColNum
End Property

'@Override
Public Property Get DataType_DataSheet() As Worksheet
    Set DataType_DataSheet = Me.ListSheet
End Property

Public Property Get DataType_DataRange() As Range
    Set DataType_DataRange = Me.ListRange
End Property


'----------Setter----------
Public Property Let ListSheet(ByVal new_Sheet As Worksheet)
    Set my_list_Sheet = new_Sheet
End Property

Public Property Let ListRange(ByVal new_Range As Range)
    Set my_list_Range = new_Range
End Property

Public Property Let ListHeaderRowNum(ByVal new_HeaderRowNum As Long)
    my_list_HeaderRowNum = new_HeaderRowNum
End Property

Public Property Let ListHeaderColNum(ByVal new_HeaderColNum As Long)
    my_list_HeaderColNum = new_HeaderColNum
End Property

'@Override
Public Property Let DataType_DataSheet(ByVal new_Sheet As Worksheet)
    Set DataType_DataSheet = new_Sheet
End Property

Public Property Let DataType_DataRange(ByVal new_Range As Range)
    Set DataType_DataRange = new_Range
End Property


'----------関数----------
Public Function ListConfig(ByVal args As String) As String
    ListConfig = my_list_cfg(args)
End Function

Public Function ListConfigs() As Object
    Set ListConfigs = my_list_cfg
End Function

Public Function ListRow(ByVal rowObject, Optional ByVal headerColObject = "", _
                                            Optional ByVal headerFlag As Boolean = False) As Range
    Dim colNum As Long: colNum = my_list_HeaderColNum
    If headerColObject <> "" Then colNum = Me.ListColumnNum(headerColObject)
    Set ListRow = RowSelect(my_list_Range, rowObject, colNum)
    If headerFlag = False Then Set ListRow = Except2(ListRow, my_list_HeaderCol)
End Function

Public Function ListColumn(ByVal colObject, Optional ByVal headerRowObject = "", _
                                                Optional ByVal headerFlag As Boolean = False) As Range
    Dim rowNum As Long: rowNum = my_list_HeaderRowNum
    If headerRowObject <> "" Then rowNum = Me.ListRowNum(headerRowObject)
    Set ListColumn = ColSelect(my_list_Range, colObject, rowNum)
    If headerFlag = False Then Set ListColumn = Except2(ListColumn, my_list_HeaderRow)
End Function

Public Function ListCells(ByVal rowObject, ByVal colObject) As Range
    Set ListCells = RangeSelect(my_list_Range, rowObject, colObject, my_list_HeaderColNum, my_list_HeaderRowNum)
End Function

Public Function ListHeaderRow() As Range
    Set ListHeaderRow = RowSelect(my_list_Range, my_list_HeaderRowNum)
End Function

Public Function ListHeaderCol() As Range
    Set ListHeaderCol = ColSelect(my_list_Range, my_list_HeaderColNum)
End Function

Public Function ListHeaderRowName(ByVal rowObject) As String
    ListHeaderRowName = Me.ListCells(rowObject, my_list_HeaderColNum).Value
End Function

Public Function ListHeaderColName(ByVal colObject) As String
    ListHeaderColName = Me.ListCells(my_list_HeaderRowNum, colObject).Value
End Function

Public Function ListRowNum(ByVal rowObject) As Long
    ListRowNum = Me.ListRow(rowObject).Row
End Function

Public Function ListColumnNum(ByVal colObject) As Long
    ListColumnNum = Me.ListColumn(colObject).Column
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

'@Override
Public Function DataType_DataConfig(ByVal args As String) As String
    DataType_DataConfig = Me.ListConfig(args)
End Function

Public Function DataType_DataConfigs() As Object
    Set DataType_DataConfigs = Me.ListConfigs
End Function


'----------サブルーチン----------
Public Sub ListBeforeDataSet()
    Call Me.ListRowsDelete
End Sub

Public Sub ListAfterDataSet()
    With Me
        Call .ListRowsCopyAndPaste(Range(.ListConfig("コピー元範囲")), .ListConfig("入力開始行数"), .ListConfig("入力終了行数"))
        Call .ListAutoFilter(.ListConfig("見出し終了行数") - .ListConfig("カラム名行数"))
    End With
End Sub

Public Sub ListRowsDelete(ByVal strRow As Long) '行の削除
    With Me.ListSheet
        Range(.Rows(strRow), .Rows(Rows.Count)).Delete
    End With
End Sub

Public Sub ListRowsCopyAndPaste(ByVal cpSrcRng As Range, ByVal strRow As Integer, ByVal endRow As Integer)  '範囲のコピぺ
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

Public Sub ListAutoFilter(Optional ByVal offsetRow As Integer = 0)  'オートフィルタ
    Me.ListRange.Offset(offsetRow).Resize(Me.ListRange.Rows.Count + offsetRow).AutoFilter
End Sub
