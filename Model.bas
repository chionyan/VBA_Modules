Attribute VB_Name = "Model"
Option Explicit
Dim wk_instance
'------------------------------�N��������------------------------------
Sub Sheet_Initialize()
    '�C���X�^���X�Z�b�g
    Call InstanceSet(STY010_V_flag:=True, OPINIO_V_flg:=True)
    If ActiveSheet.Name = "STY010_V" Then Set wk_instance = STY010_V_card
    If ActiveSheet.Name = "OPINIO_V" Then Set wk_instance = OPINIO_V_list
    
    Call BeforeDataSet
    With wk_instance
    
        '----------���X�g�^----------
        If TypeName(wk_instance) = "ListType" Then
            Call .ListRowsDelete(.ListConfig("���͊J�n�s��"))   '�s�폜
            Call .ListRowsCopyAndPaste(Range(.ListConfig("�R�s�[���͈�")), .ListConfig("���͊J�n�s��"), .ListConfig("���͏I���s��"))    '�s�R�s�y
            Call .ListAutoFilter(.ListConfig("���o���I���s��") - .ListConfig("�J�������s��"))   '�I�[�g�t�B���^�[
            Range(.ListConfig("�J�[�\���ʒu")).Select   '�J�[�\���ʒu�Z�b�g
            
        '----------�J�[�h�^----------
        ElseIf TypeName(wk_instance) = "CardType" Then
            Call .CardClearContents '�S�Z���N���A
            Call .CardStyleSetting  '�w�i�F������
            Range(.CardConfig("�J�[�\���ʒu")).Select   '�J�[�\���ʒu�Z�b�g
        End If
        
    End With
    Call AfterDataSet
End Sub

'------------------------------�{�^������------------------------------

'------------------------------�V�[�g�C�x���g------------------------------
Sub OPINIO_V_Worksheet_Change(ByVal Target As Range)
    With OPINIO_V_list
        Select Case Target.Column
            Case RangeMinColNum(.ListCells(.ListConfig("�J�������s��"), "^(Q|A).*$")) To _
                    RangeMaxColNum(.ListCells(.ListConfig("�J�������s��"), "^(Q|A).*$"))
            Call OPINIO_V_DateInput(Target)
        End Select
    End With
End Sub

'���t�������́iWorksheet_Change�j
Sub OPINIO_V_DateInput(ByVal Target As Range)
    Dim tempStr As String
    Dim dateRng As Range
    Dim ckRng As Range
    With OPINIO_V_list
        tempStr = .ListCells(.ListConfig("�J�������s��"), Target.Column).Value
        Set dateRng = .ListCells(Target.Row, "^(" & Left(tempStr, InStr(tempStr, "_") - 1) & ").*(_DATE)$")
        Set ckRng = .ListCells(Target.Row, "^(" & Left(tempStr, InStr(tempStr, "_") - 1) & ").*(_NAME|_SHEET|_SUBJECT)$")
        Call BeforeDataSet
        If False Then
        ElseIf WkstCountA(ckRng) > 0 Then dateRng.Value = Format(Date, "yy/m/d")
        ElseIf WkstCountA(ckRng) <= 0 Then dateRng.Value = ""
        End If
        Call AfterDataSet
    End With
End Sub
