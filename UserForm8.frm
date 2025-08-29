VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "撮影日時を表示します。"
   ClientHeight    =   2895
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4968
   OleObjectBlob   =   "UserForm8.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox4_Click()
' 曜日表示チェックボックス
    ' 曜日表示チェックボックスがＯＮなら
    If CheckBox4.Value = True Then
        ' 曜日表示言語ラジオボタンを有効にする
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    ' 曜日表示チェックボックスがＯＦＦなら
    Else
        ' 曜日表示言語ラジオボタンを無効にする
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
    End If
End Sub

Private Sub CommandButton1_Click()
' 決定ボタン
    ' フォントサイズが数値でない場合
    If Not IsNumeric(ComboBox1.Value) Then
        MsgBox "フォントサイズは、数値を入力してください。", vbOKOnly + vbExclamation, "お知らせ"
        Exit Sub
    ElseIf ComboBox1.Value < 0 Then
        MsgBox "フォントサイズは、ゼロ以上の数値を入力してください。", vbOKOnly + vbExclamation, "お知らせ"
        Exit Sub
    End If
    ' フォントサイズを記憶
    Range(DateFontSize).Value = ComboBox1.Value
    ' 太字にするフラグを記憶
    If CheckBox1.Value = True Then
        Range(DateFontBold).Value = 1
    Else
        Range(DateFontBold).Value = ""
    End If
    ' 撮影日時の時刻も表示させるか日にちだけを表示させるかのフラグを記憶
    If CheckBox2.Value = True Then
        Range(PictureDateType).Value = 1
    Else
        Range(PictureDateType).Value = ""
    End If
    ' フォントの色が数値で無い場合
    If (Not IsNumeric(TextBox1.Value)) Or (Not IsNumeric(TextBox2.Value)) Or (Not IsNumeric(TextBox3.Value)) Then
        MsgBox "フォントの色は、数値を入力してください。", vbOKOnly + vbExclamation, "お知らせ"
        Exit Sub
    End If
    ' フォントの色が０～２５５の値でない場合
    If TextBox1.Value < 0 Or TextBox1.Value > 255 Or _
        TextBox2.Value < 0 Or TextBox2.Value > 255 Or _
        TextBox3.Value < 0 Or TextBox3.Value > 255 Then
        MsgBox "フォントの色は、0 ～ 255 の数値で指定してください。", vbOKOnly + vbExclamation, "お知らせ"
        Exit Sub
    End If
    ' フォントの色を記憶
    Range(DateFontColorR).Value = Int(TextBox1.Value)
    Range(DateFontColorG).Value = Int(TextBox2.Value)
    Range(DateFontColorB).Value = Int(TextBox3.Value)
    ' テキストボックスの表示位置オフセットが数値でない場合
    If (Not IsNumeric(TextBox4.Value)) Or (Not IsNumeric(TextBox5.Value)) Then
        MsgBox "表示位置のオフセットは、数値を入力してください。", vbOKOnly + vbExclamation, "お知らせ"
        Exit Sub
    End If
    ' テキストボックスの表示位置オフセットの値を記憶
    Range(DateXOffset).Value = TextBox4.Value / XUnit
    Range(DateYOffset).Value = TextBox5.Value / YUnit
    ' 撮影日区切り記号「．」の使用フラグを記憶
    If CheckBox3.Value = True Then
        Range(DateSeparator).Value = 1
    Else
        Range(DateSeparator).Value = ""
    End If
    ' 曜日表示の値を記憶
    If CheckBox4.Value Then
        Range(WeekDisp).Value = 1
    Else
        Range(WeekDisp).Value = ""
    End If
    ' 曜日表示言語の値を記憶
    If OptionButton2.Value Then
        Range(WeekLang).Value = 1
    Else
        Range(WeekLang).Value = ""
    End If
    ' 画面表示の更新をしないようにする
    Application.ScreenUpdating = False
    ' 撮影日をいったん消去
    Call PictureDateOFF
    ' 撮影日を表示させる
    Call PictureDateON
    ' 画面表示の更新を許可する
    Application.ScreenUpdating = True
    ' ユーザーフォームを解放
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' キャンセルボタン
    ' ユーザーフォームを解放
    Unload Me
End Sub

Private Sub TextBox1_Change()
' テキストボックスの値が変わった場合
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ラベルの色を設定
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
            ' ユーザーフォームを再描画
            Me.Repaint
        End If
    End If
End Sub

Private Sub TextBox2_Change()
' テキストボックスの値が変わった場合
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ラベルの色を設定
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
            ' ユーザーフォームを再描画
            Me.Repaint
        End If
    End If
End Sub

Private Sub TextBox3_Change()
' テキストボックスの値が変わった場合
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ラベルの色を設定
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
            ' ユーザーフォームを再描画
            Me.Repaint
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期化処理
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' 各コントロールを無効にする
        CommandButton1.Enabled = False
        ComboBox1.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
        Exit Sub
    End If
    ' コンボボックスのリストを作成
    With ComboBox1
        .AddItem 10, 0
        .AddItem 12, 1
        .AddItem 14, 2
        .AddItem 16, 3
    End With
    ' コンボボックスの値を設定
    ComboBox1.Value = Range(DateFontSize).Value
    ' コンボボックスの値がリストにない場合
    If ComboBox1.MatchFound = False Then
        ' リストに値を追加
        ComboBox1.AddItem Range(DateFontSize).Value, 4
    End If
    ' チェックボックスの初期設定（太字の設定）
    If Range(DateFontBold).Value <> 0 Then
        CheckBox1.Value = True
    Else
        CheckBox1.Value = False
    End If
    ' テキストボックスの初期設定（フォントの色）
    TextBox1.Value = Range(DateFontColorR).Value
    If TextBox1.Value = "" Then
        TextBox1.Value = 0
    End If
    TextBox2.Value = Range(DateFontColorG).Value
    If TextBox2.Value = "" Then
        TextBox2.Value = 0
    End If
    TextBox3.Value = Range(DateFontColorB).Value
    If TextBox3.Value = "" Then
        TextBox3.Value = 0
    End If
    ' ラベルの色を設定
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ラベルの色を設定
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
        End If
    End If
    ' チェックボックスの初期設定（撮影日のみ表示）
    If Range(PictureDateType).Value <> 0 Then
        CheckBox2.Value = True
    Else
        CheckBox2.Value = False
    End If
    ' テキストボックスの初期設定（表示位置オフセット）
    TextBox4.Value = Round(Range(DateXOffset).Value * XUnit, 1)
    TextBox5.Value = Round(Range(DateYOffset).Value * YUnit, 1)
    ' チェックボックスの初期設定（撮影日区切りの設定）
    If Range(DateSeparator).Value <> 0 Then
        CheckBox3.Value = True
    Else
        CheckBox3.Value = False
    End If
    ' オプションボタンの初期設定（曜日表示の言語）
    If Range(WeekLang).Value <> 0 Then
        OptionButton2.Value = True
    Else
        OptionButton1.Value = True
    End If
    ' チェックボックスの初期設定（曜日表示の設定）
    If Range(WeekDisp).Value <> 0 Then
        CheckBox4.Value = True
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    Else
        CheckBox4.Value = False
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
    End If
End Sub

