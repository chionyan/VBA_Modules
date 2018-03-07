Attribute VB_Name = "M06_AutoFilterFunction"
'Option Explicit
Dim filDics As Object
Sub AutoFilterInitialize(ByVal argRng As Range)
    With argRng
        Call AutoFilterOff(argRng)
        Call AutoFilterOn(argRng)
    End With
End Sub
Sub AutoFilterOff(ByVal argRng As Range)
    With argRng
        If argRng.Parent.AutoFilterMode Then .AutoFilter
    End With
End Sub

Sub AutoFilterOn(ByVal argRng As Range)
    On Error GoTo errProc
    Set filDics = filDics_Attr
    With argRng
        If AutoFilterSaveCheck(filDics) = False Then
            If Not .Parent.AutoFilterMode Then .AutoFilter
        Else
            For Each Field In filDics.keys
                Criteria1 = filDics(Field)("Criteria1")
                Operator = filDics(Field)("Operator")
                Criteria2 = filDics(Field)("Criteria2")
                
                .AutoFilter Field:=Field
                If TypeName(Operator) = "Empty" Then
                    .AutoFilter Field:=Field, Criteria1:=Criteria1
                ElseIf TypeName(Criteria2) = "Empty" Then
                    .AutoFilter Field:=Field, Criteria1:=Criteria1, Operator:=Operator
                Else
                    .AutoFilter Field:=Field, Criteria1:=Criteria1, Operator:=Operator, Criteria2:=Criteria2
                End If
            Next
        End If
    End With
    Exit Sub
errProc:
    If Not argRng.Parent.AutoFilterMode Then argRng.AutoFilter
End Sub

Sub AutoFilterSave()
    Set filDics = CreateObject("Scripting.Dictionary")
    With ActiveSheet
        If .AutoFilterMode Then
            With .AutoFilter.Filters
                For i = 1 To .Count
                    If .Item(i).On Then
                        Dim filDic As Object
                        Set filDic = CreateObject("Scripting.Dictionary")
                        Criteria1 = .Item(i).Criteria1
                        Call filDic.Add("Criteria1", Criteria1)
                        If .Item(i).Operator <> 0 Then
                            On Error Resume Next
                            Operator = .Item(i).Operator
                            If Operator >= 3 And Operator <= 6 Then Operator = 1
                            Criteria2 = .Item(i).Criteria2
                            Call filDic.Add("Operator", Operator)
                            Call filDic.Add("Criteria2", Criteria2)
                        End If
                        Call filDics.Add(i, filDic)
                    End If
                Next
            End With
        End If
    End With
    Set filDics = filDics_Attr(filDics)
End Sub

Sub AutoFilterNothing()
    Call AutoFilterSave
    For Each Field In filDics.keys
        On Error Resume Next
        Call filDics(Field).Remove("Criteria1")
        Call filDics(Field).Remove("Operator")
        Call filDics(Field).Remove("Criteria2")
    Next
    Set filDics = filDics_Attr(filDics)
End Sub

Function AutoFilterSaveCheck(ByVal filDics As Object) As Boolean
    AutoFilterSaveCheck = False
    For Each Field In filDics.keys
        If filDics(Field)("Criteria1") <> "" Then
            AutoFilterSaveCheck = True: Exit Function
        End If
    Next
End Function


Private Function filDics_Attr(Optional ByVal argDic As Object = Nothing) As Object
    Static filDics As Object
    If Not argDic Is Nothing Then Set filDics = argDic
    Set filDics_Attr = filDics
End Function
