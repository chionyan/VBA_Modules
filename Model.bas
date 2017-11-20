Attribute VB_Name = "Model"
Option Explicit
Dim wk_instance
'------------------------------起動時処理------------------------------
Sub Sheet_Initialize()
    'インスタンスセット
    Call InstanceSet(STY010_V_flag:=True, OPINIO_V_flg:=True)
    If ActiveSheet.Name = "STY010_V" Then Set wk_instance = STY010_V_card
    If ActiveSheet.Name = "OPINIO_V" Then Set wk_instance = OPINIO_V_list
    
    Call BeforeDataSet
    With wk_instance
    
        '----------リスト型----------
        If TypeName(wk_instance) = "ListType" Then
            Call .ListRowsDelete(.ListConfig("入力開始行数"))   '行削除
            Call .ListRowsCopyAndPaste(Range(.ListConfig("コピー元範囲")), .ListConfig("入力開始行数"), .ListConfig("入力終了行数"))    '行コピペ
            Call .ListAutoFilter(.ListConfig("見出し終了行数") - .ListConfig("カラム名行数"))   'オートフィルター
            Range(.ListConfig("カーソル位置")).Select   'カーソル位置セット
            
        '----------カード型----------
        ElseIf TypeName(wk_instance) = "CardType" Then
            Call .CardClearContents '全セルクリア
            Call .CardStyleSetting  '背景色初期化
            Range(.CardConfig("カーソル位置")).Select   'カーソル位置セット
        End If
        
    End With
    Call AfterDataSet
End Sub

'------------------------------ボタン処理------------------------------

'------------------------------シートイベント------------------------------
Sub OPINIO_V_Worksheet_Change(ByVal Target As Range)
    With OPINIO_V_list
        Select Case Target.Column
            Case RangeMinColNum(.ListCells(.ListConfig("カラム名行数"), "^(Q|A).*$")) To _
                    RangeMaxColNum(.ListCells(.ListConfig("カラム名行数"), "^(Q|A).*$"))
            Call OPINIO_V_DateInput(Target)
        End Select
    End With
End Sub

'日付自動入力（Worksheet_Change）
Sub OPINIO_V_DateInput(ByVal Target As Range)
    Dim tempStr As String
    Dim dateRng As Range
    Dim ckRng As Range
    With OPINIO_V_list
        tempStr = .ListCells(.ListConfig("カラム名行数"), Target.Column).Value
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
