Attribute VB_Name = "TableClassExample"
Option Explicit
Dim tb1 As Table
Dim tb2 As Table

Sub TableClassExample()
    
    On Error Resume Next
    
    '�͈͐ݒ�
    '�e�[�u�����쐬����Range���w�肵�Ă��������B
    Dim rng1 As Range: Set rng1 = Range("A5:H16")
    Dim rng2 As Range: Set rng2 = Range("K14:R25")
    
     '�C���X�^���X����
     '��قǎw�肵��Range���2�����ɂ��Ă��������B
    Set tb1 = Init(New Table, rng1)
    Set tb2 = Init(New Table, rng2)
    
    ' ���\�b�h�i���낢��j
    Debug.Print tb1.TableRange.Address
    Debug.Print tb1.TableColumn("��").Address

    '�ȉ���͓��`
    Debug.Print tb1.TableRow("��").Address
    Debug.Print tb1.TableRow(7).Address

    '�ȉ���͓��`
    Debug.Print tb1.TableCells("��", "��").Value
    Debug.Print tb1.TableCells(7, "��").Value

    'sampleRange2������
    Debug.Print tb2.TableRange.Address
    Debug.Print tb2.TableRow("h").Address
    
End Sub

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

