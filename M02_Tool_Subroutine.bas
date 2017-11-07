Attribute VB_Name = "M02_Tool_Subroutine"
'******************************************************************************************
' Procedure名：最後のセルを減らすためのマクロ
' 機能概要　  ：選択範囲を新規シートにコピー、成形する。
'               ※シートイベントはコピーされないので手動でコピーしてください
'******************************************************************************************
Sub 最後のセルを減らすためのマクロ()
    If TypeName(Selection) = "Range" Then
        '選択範囲を新規シートにコピー
        Selection.Copy
        Dim srcSheetName As String: srcSheetName = ActiveSheet.Name
        Dim addSheetName As String: addSheetName = InputBox("追加するシートの名称を入力してください。")
        If addSheetName <> "" And addSheetName <> srcSheetName Then
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = addSheetName
            Sheets(addSheetName).Paste
            Application.ScreenUpdating = False
            
            '列幅成形
            Dim i As Long
            For i = 1 To Selection.Cells(Selection.Count).Column
                Sheets(addSheetName).Rows(i).RowHeight = Sheets(srcSheetName).Rows(i).RowHeight
                Sheets(addSheetName).Columns(i).ColumnWidth = Sheets(srcSheetName).Columns(i).ColumnWidth
            Next
            
            '行幅成形
            Dim j As Long
            For j = 1 To Selection.Cells(Selection.Count).Row
                Sheets(addSheetName).Rows(j).RowHeight = Sheets(srcSheetName).Rows(j).RowHeight
                Sheets(addSheetName).Columns(j).ColumnWidth = Sheets(srcSheetName).Columns(j).ColumnWidth
            Next
            Application.ScreenUpdating = True
        Else
            addSheetName = InputBox("追加するシートの名称を入力してください。")
        End If
    Else
        MsgBox "セルを選択してください。"
    End If
End Sub

'******************************************************************************************
' Procedure名：背景色取得
' 機能概要　  ：アクティブセルの背景色をRGB形式で表示する。
'******************************************************************************************
Sub 背景色取得()
    myColor = ActiveCell.Interior.Color
    
    myR = myColor Mod 256
    myG = Int(myColor / 256) Mod 256
    myB = Int(myColor / 256 / 256)
    
    MsgBox "選択したセルのRGBは、" & Chr(10) & _
            "R:" & myR & Chr(10) & _
            "G:" & myG & Chr(10) & _
            "B:" & myB & Chr(10) & _
            "です。"
End Sub

