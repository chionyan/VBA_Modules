Attribute VB_Name = "M02_Common_Subroutine"
'3�b�����҂��Ă��idbsc�ɓ����}�N���̍Ō�ɓ����K�v����j
Sub Pause()
    Application.Wait Now + TimeValue("0:00:03")
End Sub

'�u�b�N���̑S�Ẵ��[�N�V�[�g�̃X�N���[���o�[�ƃA�N�e�B�u�Z��������������
'������True��ݒ肷��ƁA�g���E���o����\������B�i�f�t�H���g��False�j
Sub ����_��ʐ���(Optional dispVisibled As Boolean = False)
    Call BeforeDataSet
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
    Call AfterDataSet
End Sub

'��ʕ`��X�V��~�E�C�x���g��~�E�V�[�g�ی�ݒ�
Sub BeforeDataSet()
    Call GP_Stop_SCUPD: Call EventStop
'    Call SheetProtect
End Sub

'��ʕ`��X�V���A�E�C�x���g�ĊJ�E�V�[�g�ی����
Sub AfterDataSet()
'    Call SheetUnprotect
    Call EventStart: Call GP_Start_SCUPD
End Sub

'�C�x���g��~
Sub EventStop()
    Application.EnableEvents = False
End Sub

'�C�x���g�ĊJ
Sub EventStart()
    Application.EnableEvents = True
End Sub

'��ʕ`��X�V��~
Sub GP_Stop_SCUPD()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

'��ʕ`��X�V���A
Sub GP_Start_SCUPD()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'�V�[�g�ی�ݒ�(�Z���I���E�I�[�g�t�B���^�[�\)
Sub SheetProtect()
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
End Sub

'�V�[�g�ی����
Sub SheetUnprotect()
    ActiveSheet.Unprotect
End Sub

'�A�N�e�B�u�V�[�g�ۑ�
Sub ActiveSheet_Save()

    Dim filePath        As Variant
    Dim path            As String
    Dim WSH             As Variant
    Dim ret             As Byte
    Dim SaveDate As String        '���t�i�t�@�C�����Ɏg�p�j
    Dim SaveTime As String        '���ԁi�t�@�C�����Ɏg�p�j
    Dim BkName As String

    Call GP_Stop_SCUPD  '����������������Ȃ�
    Call EventStop    '�C�x���g����

    Set WSH = CreateObject("WScript.Shell")
 
    ret = MsgBox(ActiveSheet.Name & "���G�N�Z���ۑ����܂��B", vbOKCancel + vbQuestion, "dbSheetClient")

    '�L�����Z������
    If ret = vbCancel Then
        MsgBox "�L�����Z�����܂����B", vbExclamation, "dbSheetClient"
        Set WSH = Nothing
        EventStart     '�C�x���g�L��
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
                EventStart     '�C�x���g�L��
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
    Call EventStart
    Call GP_Start_SCUPD
    
End Sub

'���̃V�[�g�i�u�b�N�j�Ƀ����N�����v�Z����l�ɕϊ�����
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
            Set Target = Union2(Target, FoundCell)
        End If
    Loop
    
    Target.Select
    
    For Each c In Selection
        c.Value = c.Value
    Next
    
End Sub

