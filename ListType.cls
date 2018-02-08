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
    
    If Not p(1) Is Nothing Then my_list_HeaderRowNum = p(1): Set my_list_Range = Union2(my_list_Range, RowSelect(my_list_Range, my_list_HeaderRowNum))
    If Not p(2) Is Nothing Then my_list_HeaderColNum = p(2): Set my_list_Range = Union2(my_list_Range, ColSelect(my_list_Range, my_list_HeaderColNum))
    
    Set my_list_cfg = CreateObject("Scripting.Dictionary")
    Set dbscsetRng = BottomRightExtention(Sheets("dbscset").Range("A1"))
    For Each rng In RangeSelect(dbscsetRng, "(^行$)|(^列$)|(^範囲$)|(^その他$)", "dbscset見出し")
        my_list_cfg.Add rng.Value, RangeSelect(dbscsetRng, "^" & rng.Value & "$", "^" & my_list_Sheet.Name & "$", 2).Value
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
    Set my_list_Range = Union2(my_list_Range, Me.ListHeaderRow)
End Property

Public Property Let ListHeaderColNum(ByVal new_HeaderColNum As Long)
    my_list_HeaderColNum = new_HeaderColNum
    Set my_list_Range = Union2(my_list_Range, Me.ListHeaderCol)
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
    On Error GoTo errProc
    ListConfig = my_list_cfg(args)
    Exit Function
errProc:
    ListConfig = ""
End Function

Public Function ListConfigs() As Object
    On Error GoTo errProc
    Set ListConfigs = my_list_cfg
    Exit Function
errProc:
    Set ListConfigs = Nothing
End Function

Public Function ListDataRow() As Range
    On Error GoTo errProc
    Set ListDataRow = Me.ListRow(Me.ListConfig("データ開始行数") & ":" & Me.ListConfig("データ終了行数"))
    Exit Function
errProc:
    Set ListDataRow = Nothing
End Function

Public Function ListDataColumn(ByVal colObject, Optional ByVal headerRowObject = "") As Range
    On Error GoTo errProc
    Set ListDataColumn = Intersect2(Me.ListColumn(colObject, headerRowObject), Me.ListDataRow)
    Exit Function
errProc:
    Set ListDataColumn = Nothing
End Function

Public Function ListDataCells(ByVal rowObject, ByVal colObject) As Range
    On Error GoTo errProc
    Set ListDataCells = Intersect2(Me.ListDataColumn(colObject), Me.ListRow(rowObject))
    Exit Function
errProc:
    Set ListDataCells = Nothing
End Function

Public Function ListRow(ByVal rowObject, Optional ByVal headerColObject = "") As Range
    On Error GoTo errProc
    Dim colNum As Long: colNum = my_list_HeaderColNum
    If headerColObject <> "" Then colNum = Me.ListColumnNum(headerColObject)
    Set ListRow = RowSelect(my_list_Range, rowObject, colNum)
    Exit Function
errProc:
    Set ListRow = Nothing
End Function

Public Function ListColumn(ByVal colObject, Optional ByVal headerRowObject = "") As Range
    On Error GoTo errProc
    Dim rowNum As Long: rowNum = my_list_HeaderRowNum
    If headerRowObject <> "" Then rowNum = Me.ListRowNum(headerRowObject)
    Set ListColumn = ColSelect(my_list_Range, colObject, rowNum)
    Exit Function
errProc:
    Set ListColumn = Nothing
End Function

Public Function ListCells(ByVal rowObject, ByVal colObject) As Range
    On Error GoTo errProc
    Set ListCells = RangeSelect(my_list_Range, rowObject, colObject, my_list_HeaderColNum, my_list_HeaderRowNum)
    Exit Function
errProc:
    Set ListCells = Nothing
End Function

Public Function ListHeaderRow() As Range
    On Error GoTo errProc
    Set ListHeaderRow = RowSelect(my_list_Range, my_list_HeaderRowNum)
    Exit Function
errProc:
    Set ListHeaderRow = Nothing
End Function

Public Function ListHeaderCol() As Range
    On Error GoTo errProc
    Set ListHeaderCol = ColSelect(my_list_Range, my_list_HeaderColNum)
    Exit Function
errProc:
    Set ListHeaderCol = Nothing
End Function

Public Function ListHeaderRowName(ByVal rowObject) As String
    On Error GoTo errProc
    ListHeaderRowName = Me.ListCells(rowObject, my_list_HeaderColNum).Value
    Exit Function
errProc:
    ListHeaderRowName = ""
End Function

Public Function ListHeaderColName(ByVal colObject) As String
    On Error GoTo errProc
    ListHeaderColName = Me.ListCells(my_list_HeaderRowNum, colObject).Value
    Exit Function
errProc:
    ListHeaderColName = ""
End Function

Public Function ListRowNum(ByVal rowObject) As Long
    On Error GoTo errProc
    ListRowNum = Me.ListRow(rowObject).Row
    Exit Function
errProc:
    ListRowNum = 0
End Function

Public Function ListColumnNum(ByVal colObject) As Long
    On Error GoTo errProc
    ListColumnNum = Me.ListColumn(colObject).Column
    Exit Function
errProc:
    ListColumnNum = 0
End Function

Public Function ListMinColNum() As Long
    On Error GoTo errProc
    ListMinColNum = RangeMinColNum(my_list_Range)
    Exit Function
errProc:
    ListMinColNum = 0
End Function

Public Function ListMaxColNum() As Long
    On Error GoTo errProc
    ListMaxColNum = RangeMaxColNum(my_list_Range)
    Exit Function
errProc:
    ListMaxColNum = 0
End Function

Public Function ListMinRowNum() As Long
    On Error GoTo errProc
    ListMinRowNum = RangeMinRowNum(my_list_Range)
    Exit Function
errProc:
    ListMinRowNum = 0
End Function

Public Function ListMaxRowNum() As Long
    On Error GoTo errProc
    ListMaxRowNum = RangeMaxRowNum(my_list_Range)
    Exit Function
errProc:
    ListMaxRowNum = 0
End Function

'@Override
Public Function DataType_DataConfig(ByVal args As String) As String
    DataType_DataConfig = Me.ListConfig(args)
End Function

Public Function DataType_DataConfigs() As Object
    Set DataType_DataConfigs = Me.ListConfigs
End Function
