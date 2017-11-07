Attribute VB_Name = "M01_Common_Subroutine"
'******************************************************************************************
' Procedure���F����_��ʐ���
' �@�\�T�v�@  �F�u�b�N���̑S�Ẵ��[�N�V�[�g�̃X�N���[���o�[�ƃA�N�e�B�u�Z��������������B
'               ������True��ݒ肷��ƁA�g���E���o����\������B�i�f�t�H���g��False�j
'******************************************************************************************
Sub ����_��ʐ���(Optional dispVisibled As Boolean = False)
    Application.ScreenUpdating = False
    For i = Worksheets.Count To 1 Step -1
        Worksheets(i).Select
        With ActiveWindow
            .ScrollRow = 1
            .ScrollColumn = 1
            .DisplayHeadings = dispVisibled
            .DisplayGridlines = dispVisibled
        End With
        Cells(1, 1).Select
    Next
    Application.ScreenUpdating = True
End Sub
'******************************************************************************************
' Procedure���F����_�s�R�s�[
' �@�\�T�v�@  �F����1�̃��[�N�V�[�g���̈���2�̍s������3�̍s�������4�̍s�܂ŏ����̂݃R�s�[����B
'******************************************************************************************
Sub ����_�s�R�s�[(ByVal wkSheet As Worksheet, ByVal copySourceRow As Long, _
                    ByVal dataStartRow As Long, ByVal dataEndRow As Long)
    Application.ScreenUpdating = False
        With wkSheet
            Application.EnableEvents = False
            .Rows(copySourceRow).Hidden = False
            .Rows(copySourceRow).Copy
            Dim i As Long
            For i = dataStartRow To dataEndRow
                .Rows(i).PasteSpecial Paste:=xlPasteFormats
                .Rows(i).PasteSpecial Paste:=xlPasteValidation
                For Each rng In RightExtention(Cells(copySourceRow, 1))
                    If Left(rng.FormulaR1C1, 1) = "=" Then .Cells(i, rng.Column).FormulaR1C1 = rng.FormulaR1C1
                Next
            Next
            .Rows(copySourceRow).Hidden = True
            Application.EnableEvents = True
        End With
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub
'******************************************************************************************
' Procedure���F����_�s�폜
' �@�\�T�v�@  �F����1�̃��[�N�V�[�g���̈���2�̍s�ȉ���S�č폜����B
'******************************************************************************************
Sub ����_�s�폜(ByVal wkSheet As Worksheet, ByVal dataStartRow As Long)
    wkSheet.Rows(dataStartRow & ":" & Rows.Count).Delete
End Sub
'******************************************************************************************
' Procedure���F����_�l�N���A
' �@�\�T�v�@  �F����1�̃��[�N�V�[�g���̈���2�̍s�ȉ��̒l��S�ăN���A����B
'******************************************************************************************
Sub ����_�l�N���A(ByVal wkSheet As Worksheet, ByVal dataStartRow As Long)
    wkSheet.Rows(dataStartRow & ":" & Rows.Count).ClearContents
End Sub
'******************************************************************************************
' Procedure���FEVENTSTOP
' �@�\�T�v�@  �F�C�x���g��~
'******************************************************************************************
Sub EVENTSTOP()
    Application.EnableEvents = False
End Sub

'******************************************************************************************
' Procedure���FEVENTSTART
' �@�\�T�v�@  �F�C�x���g�ĊJ
'******************************************************************************************
Sub EVENTSTART()
    Application.EnableEvents = True
End Sub
'******************************************************************************************
' Procedure���FGP_Stop_SCUPD
' �@�\�T�v�@  �F��ʕ`��X�V��~
'******************************************************************************************
Sub GP_Stop_SCUPD()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub
'******************************************************************************************
' Procedure���FGP_Start_SCUPD
' �@�\�T�v�@  �F��ʕ`��X�V���A
'******************************************************************************************
Sub GP_Start_SCUPD()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
'******************************************************************************************
' Procedure���FSheetProtect
' �@�\�T�v�@  �F�V�[�g�ی�ݒ�(�Z���I���E�I�[�g�t�B���^�[�\)
'******************************************************************************************
Sub SheetProtect()
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
End Sub
'******************************************************************************************
' Procedure���FSheetUnprotect
' �@�\�T�v�@  �F�V�[�g�ی����
'******************************************************************************************
Sub SheetUnprotect()
    ActiveSheet.Unprotect
End Sub
'******************************************************************************************
' Procedure���FActiveSheet_Save
' �@�\�T�v�@  �F�A�N�e�B�u�V�[�g�ۑ�
'******************************************************************************************
Sub ActiveSheet_Save()

    Dim filePath        As Variant
    Dim path            As String
    Dim WSH             As Variant
    Dim ret             As Byte
    Dim SaveDate As String        '���t�i�t�@�C�����Ɏg�p�j
    Dim SaveTime As String        '���ԁi�t�@�C�����Ɏg�p�j
    Dim BkName As String

    Call GP_Stop_SCUPD  '����������������Ȃ�
    Call EVENTSTOP    '�C�x���g����

    Set WSH = CreateObject("WScript.Shell")
 
    ret = MsgBox(ActiveSheet.Name & "���G�N�Z���ۑ����܂��B", vbOKCancel + vbQuestion, "dbSheetClient")

    '�L�����Z������
    If ret = vbCancel Then
        MsgBox "�L�����Z�����܂����B", vbExclamation, "dbSheetClient"
        Set WSH = Nothing
        EVENTSTART     '�C�x���g�L��
        GP_Start_SCUPD
        Exit Sub
    End If

    '�}�C�h�L�������g�w��
    path = WSH.SpecialFolders("MyDocuments") & "\"
    ChDir path

    '�_�C�A���O��\�����ۑ�����p�X���擾
    SaveDate = Replace(Date, "/", "")
    SaveTime = Replace(Time, ":", "")
    BkName = ActiveSheet.Name & "_" & SaveDate & "_" & SaveTime
    filePath = Application.GetSaveAsFilename(BkName, "ExcelBook, *.xlsx, ExcelBook, *.xls", 1)

    '�t�@�C�����w�肳��Ă���΁A�ۑ����������s
    If filePath = False Then
        MsgBox "�L�����Z�����܂����B", vbExclamation, "dbSheetClient"
    Else
    
        '// �㏑���ۑ��m�F
        If Dir(filePath) <> "" Then
        
            ret = MsgBox(ActiveSheet.Name & "�͊��ɑ��݂��܂��B" & vbCrLf & "�㏑�����܂����H", _
                                                             vbYesNo + vbExclamation, "���O��t���ĕۑ��̊m�F")
        
            '// �㏑�����Ȃ��ꍇ�͏����I��
            If ret = vbNo Then
'                Application.DisplayAlerts = True
                Set WSH = Nothing
                EVENTSTART     '�C�x���g�L��
                GP_Start_SCUPD
                Exit Sub
            End If
        End If
   
        '�V�[�g���R�s�[
        ThisWorkbook.ActiveSheet.Copy
        Application.DisplayAlerts = False       '�t�@�C������鎞�̕ۑ����܂����H�_�C�A���O�̕\����}��

        '�G�N�Z���`���ŕۑ�
        ActiveWorkbook.SaveAs filePath, xlWorkbookDefault
        
        '// �ʃu�b�N�Q�ƌv�Z����l��
        Call Conv_LinkFormula
        '�v���W�F�N�g���ޖ���l��
        For Each Header In RightExtention(ActiveWorkbook.ActiveSheet.Cells(9, "B"))
            If Header.Value = "PROJECT_CLASS_NAME" Then
                For Each c In BottomExtention(Header.Offset(1, 0))
                    c.Value = c.Value
                Next
                Exit For
            End If
        Next
        '���O�����ׂč폜����
        Dim nm As Name
        For Each nm In ActiveWorkbook.Names
            On Error Resume Next  ' �G���[�𖳎��B
            nm.Delete
        Next nm

        Application.Calculation = xlCalculationAutomatic    '�u�b�N�ۑ��O�ɍČv�Z�������ɂ���
        ActiveSheet.Unprotect
        Workbooks(ActiveWorkbook.Name).Save     '�㏑�ۑ�
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
    End If

    Set WSH = Nothing
    Call EVENTSTART
    Call GP_Start_SCUPD
    
End Sub
'******************************************************************************************
' Procedure���FConv_LinkFormula
' �@�\�T�v�@  �F���̃V�[�g�i�u�b�N�j�Ƀ����N�����v�Z����l�ɕϊ�����
'******************************************************************************************
Sub Conv_LinkFormula()

    '�ق��̃V�[�g���Q�Ƃ��Ă��鐔����l�ɕϊ�
    
    Dim FoundCell As Range
    Dim FirstCell As Range
    Dim Target As Range
    Dim c As Range
    
    Set FoundCell = Cells.Find(What:="[")
    
    If FoundCell Is Nothing Then
        Exit Sub
    Else
        Set FirstCell = FoundCell
        Set Target = FoundCell
    End If
    
    Do
        Set FoundCell = Cells.FindNext(FoundCell)
        If FoundCell.Address = FirstCell.Address Then
            Exit Do
        Else
            Set Target = Union(Target, FoundCell)
        End If
    Loop
    
    Target.Select
    
    For Each c In Selection
        c.Value = c.Value
    Next
    
End Sub

