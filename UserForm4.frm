VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "処理を選択してください。"
   ClientHeight    =   6610
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   4572
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' リストの行数の最大値を定数として設定する
Const myMaxRowCount As Long = 65536
' リストのシート名を定数として設定する
Const myItemSheetName As String = "用語集"

Private Sub CheckBox2_Click()
' 折り返して表示チェックボックス
    If CheckBox2.Value Then
        CheckBox1.Enabled = False
    Else
        CheckBox1.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
' コピーボタン
    Dim myTopRow As Long, myRowsCount As Long
    ' 選択範囲のうち１列目だけを選択する
    Selection.Columns(1).Select
    ' 選択範囲の開始行を求める
    myTopRow = Selection.Row
    ' 選択範囲の行数を求める
    myRowsCount = Selection.Rows.Count
    ' 選択範囲の行数が規定値を超える場合
    If (myTopRow Mod 11) + myRowsCount > 10 Then
        ' 選択範囲の行数を制限する
        myRowsCount = 11 - (myTopRow Mod 11)
    End If
    ' 選択範囲を制限する
    Selection.Resize(myRowsCount).Select
    ' 選択範囲をコピー
    Selection.Copy
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton10_Click()
' ページ追加ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' ページ追加を呼び出す（引数は追加枚数）
    Call AddPageProc(1)
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton11_Click()
' 印刷ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' 印刷プレビューを呼び出す
    UserForm1.Show vbModal
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton12_Click()
' 続きに取込ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' 一括取り込みを呼び出す
    Call GetMultiPicture
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton13_Click()
' 適用ボタン
    Dim myTopRow As Long, myRowsCount As Long
    ' 選択範囲のうち１列目だけを選択する
    Selection.Columns(1).Select
    ' 選択範囲の開始行を求める
    myTopRow = Selection.Row
    ' 選択範囲の行数を求める
    myRowsCount = Selection.Rows.Count
    ' 選択範囲の行数が規定値を超える場合
    If (myTopRow Mod 11) + myRowsCount > 10 Then
        ' 選択範囲の行数を制限する
        myRowsCount = 11 - (myTopRow Mod 11)
    End If
    ' 選択範囲を制限する
    Selection.Resize(myRowsCount).Select
    ' 縮小して全体を表示
    Selection.ShrinkToFit = CheckBox1.Value
    ' 折り返して全体を表示
    Selection.WrapText = CheckBox2.Value
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' 貼り付けボタン
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' 罫線以外を貼り付け
    Selection.PasteSpecial Paste:=xlPasteAllExceptBorders
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' キャンセルボタン
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton14_Click()
' コマ毎に番号振直しボタン
    Me.Hide
    ' コマ毎の番号付番処理を呼び出す
    Call SerialNumbering
    Unload Me
End Sub

Private Sub CommandButton15_Click()
' 写真毎に番号振直しボタン
    Me.Hide
    ' 写真毎の番号付番処理を呼び出す
    Call PictureNumbering
    Unload Me
End Sub

Private Sub CommandButton16_Click()
' ショートカット表示ボタン
    ' ショートカット表示フラグをセット
    ShortCutFlag = False
    Unload Me
End Sub

Private Sub CommandButton17_Click()
' マクロなしで保存終了ボタン
    Me.Hide
    ' マクロなしで保存終了処理を呼び出す
    Call SaveWOMacro
    Unload Me
End Sub

Private Sub CommandButton18_Click()
' 撮影日表示ボタン
    Me.Hide
    ' 撮影日時表示シーケンスを呼び出す
    Call PictureDateDispSequence
    Unload Me
End Sub

Private Sub CommandButton19_Click()
' 末尾のページ削除ボタン
    Me.Hide
    ' 末尾のページ削除処理を呼び出す
    Call DeleteLastPages
    Unload Me
End Sub


Private Sub CommandButton4_Click()
' セルに挿入ボタン（１）
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' コンボボックスのリストが選択されている場合
    If ComboBox1.ListIndex <> -1 Then
        ' コンボボックスのリストの値をセルに挿入
        Selection.Value = ComboBox1.List(ComboBox1.ListIndex)
        ' コンボボックスの先頭位置を保存
        ComboList1 = ComboBox1.ListIndex
    ' コンボボックスに入力された値がリストに無い場合
    ElseIf ComboBox1.MatchFound = False Then
        ' リストに値を追加する
        With Worksheets(myItemSheetName)
            ' リストの１行目が値なしの場合
            If .Range("A1").Value = "" Then
                ' リストの１行目にコンボボックスの値を追加
                .Range("A1").Value = ComboBox1.Value
                ' コンボボックスの先頭位置を保存
                ComboList1 = 0
            ' リストの２行目が値なしの場合
            ElseIf .Range("A2").Value = "" Then
                ' リストの２行目にコンボボックスの値を追加
                .Range("A2").Value = ComboBox1.Value
                ' コンボボックスの先頭位置を保存
                ComboList1 = 1
            ' リストの最終行が最大値未満の場合
            ElseIf .Range("A1").End(xlDown).Row < myMaxRowCount Then
                ' リストの最終行の下にコンボボックスの値を追加
                .Range("A1").End(xlDown).Offset(1).Value = ComboBox1.Value
                ' コンボボックスの先頭位置を保存
                ComboList1 = .Range("A1").End(xlDown).Row - 1
            End If
        End With
        ' コンボボックスの値をセルに挿入
        Selection.Value = ComboBox1.Value
    End If
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton5_Click()
' セルに挿入ボタン（２）
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' コンボボックスのリストが選択されている場合
    If ComboBox2.ListIndex <> -1 Then
        ' コンボボックスのリストの値をセルに挿入
        Selection.Value = ComboBox2.List(ComboBox2.ListIndex)
        ' コンボボックスの先頭位置を保存
        ComboList2 = ComboBox2.ListIndex
    ' コンボボックスに入力された値がリストにない場合
    ElseIf ComboBox2.MatchFound = False Then
        ' リストに値を追加する
        With Worksheets(myItemSheetName)
            If .Range("B1").Value = "" Then
                .Range("B1").Value = ComboBox2.Value
                ComboList2 = 0
            ElseIf .Range("B2").Value = "" Then
                .Range("B2").Value = ComboBox2.Value
                ComboList2 = 1
            ElseIf .Range("B2").End(xlDown).Row < myMaxRowCount Then
                .Range("B1").End(xlDown).Offset(1).Value = ComboBox2.Value
                ComboList2 = .Range("B1").End(xlDown).Row - 1
            End If
        End With
        ' コンボボックスの値をセルに挿入
        Selection.Value = ComboBox2.Value
    End If
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton6_Click()
' セルに挿入ボタン（３）
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' コンボボックスのリストが選択されている場合
    If ComboBox3.ListIndex <> -1 Then
        ' コンボボックスのリストの値をセルに挿入
        Selection.Value = ComboBox3.List(ComboBox3.ListIndex)
        ' コンボボックスの先頭位置を保存
        ComboList3 = ComboBox3.ListIndex
    ' コンボボックスに入力された値がリストにない場合
    ElseIf ComboBox3.MatchFound = False Then
        ' リストに値を追加する
        With Worksheets(myItemSheetName)
            If .Range("C1").Value = "" Then
                .Range("C1").Value = ComboBox3.Value
                ComboList3 = 0
            ElseIf .Range("C2").Value = "" Then
                .Range("C2").Value = ComboBox3.Value
                ComboList3 = 1
            ElseIf .Range("C1").End(xlDown).Row < myMaxRowCount Then
                .Range("C1").End(xlDown).Offset(1).Value = ComboBox3.Value
                ComboList3 = .Range("C1").End(xlDown).Row - 1
            End If
        End With
        ' コンボボックスの値をセルに挿入
        Selection.Value = ComboBox3.Value
    End If
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton7_Click()
' セルに挿入ボタン（４）
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' コンボボックスのリストが選択されている場合
    If ComboBox4.ListIndex <> -1 Then
        ' コンボボックスのリストの値をセルに挿入
        Selection.Value = ComboBox4.List(ComboBox4.ListIndex)
        ' コンボボックスの先頭位置を保存
        ComboList4 = ComboBox4.ListIndex
    ' コンボボックスに入力された値がリストにない場合
    ElseIf ComboBox4.MatchFound = False Then
        ' リストに値を追加する
        With Worksheets(myItemSheetName)
            If .Range("D1").Value = "" Then
                .Range("D1").Value = ComboBox4.Value
                ComboList4 = 0
            ElseIf .Range("D2").Value = "" Then
                .Range("D2").Value = ComboBox4.Value
                ComboList4 = 1
            ElseIf .Range("D1").End(xlDown).Row < myMaxRowCount Then
                .Range("D1").End(xlDown).Offset(1).Value = ComboBox4.Value
                ComboList4 = .Range("D1").End(xlDown).Row - 1
            End If
        End With
        ' コンボボックスの値をセルに挿入
        Selection.Value = ComboBox4.Value
    End If
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton8_Click()
' セルに挿入ボタン（５）
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' コンボボックスのリストが選択されている場合
    If ComboBox5.ListIndex <> -1 Then
        ' コンボボックスのリストの値をセルに挿入
        Selection.Value = ComboBox5.List(ComboBox5.ListIndex)
        ' コンボボックスの先頭位置を保存
        ComboList5 = ComboBox5.ListIndex
    ' コンボボックスに入力された値がリストにない場合
    ElseIf ComboBox5.MatchFound = False Then
        ' リストに値を追加する
        With Worksheets(myItemSheetName)
            If .Range("E1").Value = "" Then
                .Range("E1").Value = ComboBox5.Value
                ComboList5 = 0
            ElseIf .Range("E2").Value = "" Then
                .Range("E2").Value = ComboBox5.Value
                ComboList5 = 1
            ElseIf .Range("E1").End(xlDown).Row < myMaxRowCount Then
                .Range("E1").End(xlDown).Offset(1).Value = ComboBox5.Value
                ComboList5 = .Range("E1").End(xlDown).Row - 1
            End If
        End With
        ' コンボボックスの値をセルに挿入
        Selection.Value = ComboBox5.Value
    End If
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton9_Click()
' セルに挿入ボタン（６）
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' コンボボックスのリストが選択されている場合
    If ComboBox6.ListIndex <> -1 Then
        ' コンボボックスのリストの値をセルに挿入
        Selection.Value = ComboBox6.List(ComboBox6.ListIndex)
        ' コンボボックスの先頭位置を保存
        ComboList6 = ComboBox6.ListIndex
    ' コンボボックスに入力された値がリストにない場合
    ElseIf ComboBox6.MatchFound = False Then
        ' リストに値を追加する
        With Worksheets(myItemSheetName)
            If .Range("F1").Value = "" Then
                .Range("F1").Value = ComboBox6.Value
                ComboList6 = 0
            ElseIf .Range("F2").Value = "" Then
                .Range("F2").Value = ComboBox6.Value
                ComboList6 = 1
            ElseIf .Range("F1").End(xlDown).Row < myMaxRowCount Then
                .Range("F1").End(xlDown).Offset(1).Value = ComboBox6.Value
                ComboList6 = .Range("F1").End(xlDown).Row - 1
            End If
        End With
        ' コンボボックスの値をセルに挿入
        Selection.Value = ComboBox6.Value
    End If
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期設定
    Dim myWorksheet As Worksheet
    Dim mySheetExist As Boolean
    ' ショートカット表示フラグをリセット
    ShortCutFlag = True
    ' チェックボックスの初期設定
    CheckBox1.Value = True
    CheckBox2.Value = False
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' コマンドボタンを無効にする
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        CommandButton4.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton9.Enabled = False
        CommandButton10.Enabled = False
        CommandButton12.Enabled = False
        CommandButton13.Enabled = False
        CommandButton14.Enabled = False
        CommandButton15.Enabled = False
        CommandButton17.Enabled = False
        CommandButton18.Enabled = False
        CommandButton19.Enabled = False
        ' コンボボックス１から６を無効にする
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False
        ' 初期設定を終了
        Exit Sub
    End If
    ' クリップボードがテキストでない場合
    If Application.ClipboardFormats(1) <> xlClipboardFormatText Then
        ' 貼り付けボタンを無効にする
        CommandButton2.Enabled = False
    End If
    ' 撮影日表示ボタン名を設定
    If Range(PictureDateFlag).Value <> 0 Then
        CommandButton18.Caption = "撮影日表示ＯＦＦ(F)"
    Else
        CommandButton18.Caption = "撮影日表示ＯＮ(F)"
    End If
    ' ワークシートフラグをリセット
    mySheetExist = False
    For Each myWorksheet In Worksheets
        ' ワークシート(myItemSheetName)が存在すれば
        If myWorksheet.Name = myItemSheetName Then
            ' ワークシートフラグをセット
            mySheetExist = True
            ' くりかえし処理を抜け出す
            Exit For
        End If
    Next
    ' ワークシート(myItenSheetName)が無ければ
    If mySheetExist = False Then
        ' コマンドボタン４から９を無効にする
        CommandButton4.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton9.Enabled = False
        ' コンボボックス１から６を無効にする
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False
        ' 初期設定を終了
        Exit Sub
    End If
    With Worksheets(myItemSheetName)
        ' コンボボックス１のセル範囲を更新
        If .Range("A1") = "" Or .Range("A2") = "" Then
            ComboBox1.RowSource = .Range("A1").Address(External:=True)
        Else
            ComboBox1.RowSource = _
                    .Range("A1:" & "A" & Format(.Range("A1").End(xlDown).Row)).Address(External:=True)
        End If
        ' コンボボックス１の先頭位置を復元
        ComboBox1.ListIndex = ComboList1
        ' コンボボックス２のセル範囲を更新
        If .Range("B1") = "" Or .Range("B2") = "" Then
            ComboBox2.RowSource = .Range("B1").Address(External:=True)
        Else
            ComboBox2.RowSource = _
                    .Range("B1:" & "B" & Format(.Range("B1").End(xlDown).Row)).Address(External:=True)
        End If
        ' コンボボックス２の先頭位置を復元
        ComboBox2.ListIndex = ComboList2
        ' コンボボックス３のセル範囲を更新
        If .Range("C1") = "" Or .Range("C2") = "" Then
            ComboBox3.RowSource = .Range("C1").Address(External:=True)
        Else
            ComboBox3.RowSource = _
                    .Range("C1:" & "C" & Format(.Range("C1").End(xlDown).Row)).Address(External:=True)
        End If
        ' コンボボックス３の先頭位置を復元
        ComboBox3.ListIndex = ComboList3
        ' コンボボックス４のセル範囲を更新
        If .Range("D1") = "" Or .Range("D2") = "" Then
            ComboBox4.RowSource = .Range("D1").Address(External:=True)
        Else
            ComboBox4.RowSource = _
                    .Range("D1:" & "D" & Format(.Range("D1").End(xlDown).Row)).Address(External:=True)
        End If
        ' コンボボックス４の先頭位置を復元
        ComboBox4.ListIndex = ComboList4
        ' コンボボックス５のセル範囲を更新
        If .Range("E1") = "" Or .Range("E2") = "" Then
            ComboBox5.RowSource = .Range("E1").Address(External:=True)
        Else
            ComboBox5.RowSource = _
                    .Range("E1:" & "E" & Format(.Range("E1").End(xlDown).Row)).Address(External:=True)
        End If
        ' コンボボックス５の先頭位置を復元
        ComboBox5.ListIndex = ComboList5
        ' コンボボックス６のセル範囲を更新
        If .Range("F1") = "" Or .Range("F2") = "" Then
            ComboBox6.RowSource = .Range("F1").Address(External:=True)
        Else
            ComboBox6.RowSource = _
                    .Range("F1:" & "F" & Format(.Range("F1").End(xlDown).Row)).Address(External:=True)
        End If
        ' コンボボックス６の先頭位置を復元
        ComboBox6.ListIndex = ComboList6
    End With
End Sub
