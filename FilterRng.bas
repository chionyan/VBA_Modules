Attribute VB_Name = "ACVEXP_V"
Function FilterRng(ByVal argRng As Range, ByVal args As String) As Range

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = args
    
    Dim rng As Range
    For Each rng In argRng
        If re.Test(rng.Value) Then
            If FilterRng Is Nothing Then
                Set FilterRng = rng
            Else
                Set FilterRng = Union(FilterRng, rng)
            End If
        End If
    Next
    
End Function

