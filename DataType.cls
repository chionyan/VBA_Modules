VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------Getter----------
Public Property Get DataSheet() As Worksheet
End Property

Public Property Get DataRange() As Range
End Property

'----------Setter----------
Public Property Let DataSheet(ByVal new_Sheet As Worksheet)
End Property

Public Property Let DataRange(ByVal new_Range As Range)
End Property

'----------ä÷êî----------
Public Function DataConfig(ByVal args As String) As String
End Function

Public Function DataConfigs() As Object
End Function
