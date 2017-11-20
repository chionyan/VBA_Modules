Attribute VB_Name = "M02_Common_Subroutine"
'3秒だけ待ってやる（dbscに入れるマクロの最後に入れる必要あり）
Sub Pause()
    Application.Wait Now + TimeValue("0:00:03")
End Sub

'ブック内の全てのワークシートのスクロールバーとアクティブセルを初期化する｡
'引数にTrueを設定すると、枠線・見出しを表示する。（デフォルトはFalse）
Sub 共通_画面制御(Optional dispVisibled As Boolean = False)
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

'画面描画更新停止・イベント停止・シート保護設定
Sub BeforeDataSet()
    Call GP_Stop_SCUPD: Call EventStop
'    Call SheetProtect
End Sub

'画面描画更新復帰・イベント再開・シート保護解除
Sub AfterDataSet()
'    Call SheetUnprotect
    Call EventStart: Call GP_Start_SCUPD
End Sub

'イベント停止
Sub EventStop()
    Application.EnableEvents = False
End Sub

'イベント再開
Sub EventStart()
    Application.EnableEvents = True
End Sub

'画面描画更新停止
Sub GP_Stop_SCUPD()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

'画面描画更新復帰
Sub GP_Start_SCUPD()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'シート保護設定(セル選択・オートフィルター可能)
Sub SheetProtect()
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
End Sub

'シート保護解除
Sub SheetUnprotect()
    ActiveSheet.Unprotect
End Sub

'アクティブシート保存
Sub ActiveSheet_Save()

    Dim filePath        As Variant
    Dim path            As String
    Dim WSH             As Variant
    Dim ret             As Byte
    Dim SaveDate As String        '日付（ファイル名に使用）
    Dim SaveTime As String        '時間（ファイル名に使用）
    Dim BkName As String

    Call GP_Stop_SCUPD  '処理中動作を見せない
    Call EventStop    'イベント無効

    Set WSH = CreateObject("WScript.Shell")
 
    ret = MsgBox(ActiveSheet.Name & "をエクセル保存します。", vbOKCancel + vbQuestion, "dbSheetClient")

    'キャンセル押下
    If ret = vbCancel Then
        MsgBox "キャンセルしました。", vbExclamation, "dbSheetClient"
        Set WSH = Nothing
        EventStart     'イベント有効
        GP_Start_SCUPD
        Exit Sub
    End If

    'マイドキュメント指定
    path = WSH.SpecialFolders("MyDocuments") & "\"
    ChDir path

    'ダイアログを表示し保存するパスを取得
    SaveDate = Replace(Date, "/", "")
    SaveTime = Replace(Time, ":", "")
    BkName = ActiveSheet.Name & "_" & SaveDate & "_" & SaveTime
    filePath = Application.GetSaveAsFilename(BkName, "ExcelBook, *.xlsx, ExcelBook, *.xls", 1)

    'ファイルが指定されていれば、保存処理を実行
    If filePath = False Then
        MsgBox "キャンセルしました。", vbExclamation, "dbSheetClient"
    Else
    
        '// 上書き保存確認
        If Dir(filePath) <> "" Then
        
            ret = MsgBox(ActiveSheet.Name & "は既に存在します。" & vbCrLf & "上書きしますか？", _
                                                             vbYesNo + vbExclamation, "名前を付けて保存の確認")
        
            '// 上書きしない場合は処理終了
            If ret = vbNo Then
'                Application.DisplayAlerts = True
                Set WSH = Nothing
                EventStart     'イベント有効
                GP_Start_SCUPD
                Exit Sub
            End If
        End If
   
        'シートをコピー
        ThisWorkbook.ActiveSheet.Copy
        Application.DisplayAlerts = False       'ファイルを閉じる時の保存しますか？ダイアログの表示を抑制

        'エクセル形式で保存
        ActiveWorkbook.SaveAs filePath, xlWorkbookDefault
        
        '// 別ブック参照計算式を値化
        Call Conv_LinkFormula
        'プロジェクト分類名を値化
        For Each Header In RightExtention(ActiveWorkbook.ActiveSheet.Cells(9, "B"))
            If Header.Value = "PROJECT_CLASS_NAME" Then
                For Each c In BottomExtention(Header.Offset(1, 0))
                    c.Value = c.Value
                Next
                Exit For
            End If
        Next
        '名前をすべて削除する
        Dim nm As Name
        For Each nm In ActiveWorkbook.Names
            On Error Resume Next  ' エラーを無視。
            nm.Delete
        Next nm

        Application.Calculation = xlCalculationAutomatic    'ブック保存前に再計算を自動にする
        ActiveSheet.Unprotect
        Workbooks(ActiveWorkbook.Name).Save     '上書保存
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
    End If

    Set WSH = Nothing
    Call EventStart
    Call GP_Start_SCUPD
    
End Sub

'他のシート（ブック）にリンクを持つ計算式を値に変換する
Sub Conv_LinkFormula()

    'ほかのシートを参照している数式を値に変換
    
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

