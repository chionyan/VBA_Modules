Attribute VB_Name = "M02_Tool_Subroutine"
'******************************************************************************************
' Procedure���F�Ō�̃Z�������炷���߂̃}�N��
' �@�\�T�v�@  �F�I��͈͂�V�K�V�[�g�ɃR�s�[�A���`����B
'               ���V�[�g�C�x���g�̓R�s�[����Ȃ��̂Ŏ蓮�ŃR�s�[���Ă�������
'******************************************************************************************
Sub �Ō�̃Z�������炷���߂̃}�N��()
    If TypeName(Selection) = "Range" Then
        '�I��͈͂�V�K�V�[�g�ɃR�s�[
        Selection.Copy
        Dim srcSheetName As String: srcSheetName = ActiveSheet.Name
        Dim addSheetName As String: addSheetName = InputBox("�ǉ�����V�[�g�̖��̂���͂��Ă��������B")
        If addSheetName <> "" And addSheetName <> srcSheetName Then
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = addSheetName
            Sheets(addSheetName).Paste
            Application.ScreenUpdating = False
            
            '�񕝐��`
            Dim i As Long
            For i = 1 To Selection.Cells(Selection.Count).Column
                Sheets(addSheetName).Rows(i).RowHeight = Sheets(srcSheetName).Rows(i).RowHeight
                Sheets(addSheetName).Columns(i).ColumnWidth = Sheets(srcSheetName).Columns(i).ColumnWidth
            Next
            
            '�s�����`
            Dim j As Long
            For j = 1 To Selection.Cells(Selection.Count).Row
                Sheets(addSheetName).Rows(j).RowHeight = Sheets(srcSheetName).Rows(j).RowHeight
                Sheets(addSheetName).Columns(j).ColumnWidth = Sheets(srcSheetName).Columns(j).ColumnWidth
            Next
            Application.ScreenUpdating = True
        Else
            addSheetName = InputBox("�ǉ�����V�[�g�̖��̂���͂��Ă��������B")
        End If
    Else
        MsgBox "�Z����I�����Ă��������B"
    End If
End Sub

'******************************************************************************************
' Procedure���F�w�i�F�擾
' �@�\�T�v�@  �F�A�N�e�B�u�Z���̔w�i�F��RGB�`���ŕ\������B
'******************************************************************************************
Sub �w�i�F�擾()
    myColor = ActiveCell.Interior.Color
    
    myR = myColor Mod 256
    myG = Int(myColor / 256) Mod 256
    myB = Int(myColor / 256 / 256)
    
    MsgBox "�I�������Z����RGB�́A" & Chr(10) & _
            "R:" & myR & Chr(10) & _
            "G:" & myG & Chr(10) & _
            "B:" & myB & Chr(10) & _
            "�ł��B"
End Sub

