VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "処理を選択してください。"
   ClientHeight    =   3020
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   4800
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' ここに取込ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' ここに一括取込を呼び出す
    Call GetMultiPictureFromHere
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' ページ追加ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' ページ追加を呼び出す（引数は追加枚数）
    Call AddPageProc(1)
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' 印刷ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' 印刷プレビューを呼び出す
    UserForm1.Show vbModal
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton4_Click()
' キャンセルボタン
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton5_Click()
' 貼り付けボタン
    Dim myPictureName As String
    Dim myPicture As Shape
    Dim i As Integer, CurrentNo As Integer, myDate As String, myType As Integer
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' 画面表示の更新をしないようにする
    Application.ScreenUpdating = False
    ' 貼り付け
    ActiveSheet.Paste
    ' 写真が傾いていたら９０°刻みに補正する
    With Selection.ShapeRange
        If .Rotation < 45 Or .Rotation >= 315 Then
            .Rotation = 0
        ElseIf .Rotation >= 45 And .Rotation < 135 Then
            .Rotation = 90
        ElseIf .Rotation >= 135 And .Rotation < 225 Then
            .Rotation = 180
        ElseIf .Rotation >= 225 And .Rotation < 315 Then
            .Rotation = 270
        End If
        ' 写真を最背面に移動
        .ZOrder msoSendToBack
        ' 写真の名前を取得
        myPictureName = .Name
    End With
    ' 同じ名前のカウンタ
    i = 0
    ' すべての図に対して
    For Each myPicture In ActiveSheet.Shapes
        ' 同じ名前があれば
        If myPicture.Name = myPictureName Then
            ' カウンタを加算
            i = i + 1
            ' 同じ名前の写真があれば
            If i > 1 Then
                ' 写真を切り取る
                Selection.Cut
                ' ユーザーフォームを非表示にする
                Me.Hide
                ' メッセージを表示して貼り付けを中断させる
                MsgBox "同じ写真を貼り付けることはできません。", vbOKOnly + vbExclamation, "お知らせ"
                ' ユーザーフォームを解放する
                Application.ScreenUpdating = True
                Unload Me
                Exit Sub
            End If
        End If
    Next
    With ActiveSheet.Shapes(myPictureName)
        ' 写真を貼り付けたセルを選択
        .TopLeftCell.MergeArea.Select
        ' 「Ａ２」セルで写真が縦長で９０°または２７０°の場合に位置決めがずれる対策
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TempRowHeight
        End If
        ' 写真の縦横比を固定する
        .LockAspectRatio = msoTrue
        ' 写真の角度が９０°または２７０°の場合
        If .Rotation = 90 Or .Rotation = 270 Then
            ' 写真の幅をセルの高さにそろえる
            .Width = Int(ActiveCell.MergeArea.Height) - 4
           ' 写真の高さをセルの幅にそろえる
            If .Height > Int(ActiveCell.MergeArea.Width) - 4 Then
                .Height = Int(ActiveCell.MergeArea.Width) - 4
            End If
        ' 写真の角度が０°または１８０°の場合
        Else
            ' 写真の高さをセルの高さにそろえる
            .Height = Int(ActiveCell.MergeArea.Height) - 4
            ' 写真の幅をセルの幅にそろえる
            If .Width > Int(ActiveCell.MergeArea.Width) - 4 Then
                .Width = Int(ActiveCell.MergeArea.Width) - 4
            End If
        End If
        ' 写真の位置決め
        .Top = Selection.Top + ((Selection.Height - .Height) / 2)
        .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        ' 「Ａ２」セルで写真が縦長で９０°または２７０°の場合に位置決めがずれる対策
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TopRowHeight
        End If
    End With
    ' 撮影日データを復元
    Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = _
        Worksheets(CutDataSheet).Range(PictureNameBuffer & Format(CutDataBuffer)).Value
    Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = _
        Worksheets(CutDataSheet).Range(PictureDateBuffer & Format(CutDataBuffer)).Value
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag).Value <> 0 Then
        ' 撮影日を取得
        myDate = Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value
        ' 撮影日の表示形式フラグをエンコード
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        ' 撮影日のテキストボックスを作図
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If
    ' 画面表示の更新を許可する
    Application.ScreenUpdating = True
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton6_Click()
' 続きに取込ボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' 一括取り込みを呼び出す
    Call GetMultiPicture
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton7_Click()
' 余白コマをつめるボタン
    Me.Hide
    ' コマの削除を呼び出す
    Call DeleteBlank
    Unload Me
End Sub

Private Sub CommandButton8_Click()
' 余白コマを追加ボタン
    Me.Hide
    ' コマの追加を呼び出す
    Call AddBlank
    Unload Me
End Sub

Private Sub CommandButton9_Click()
' コマ毎に番号振直しボタン
    Me.Hide
    ' コマ毎の番号付番処理を呼び出す
    Call SerialNumbering
    Unload Me
End Sub

Private Sub CommandButton10_Click()
' 写真毎に番号振直しボタン
    Me.Hide
    ' 写真毎の番号付番処理を呼び出す
    Call PictureNumbering
    Unload Me
End Sub

Private Sub CommandButton11_Click()
' ショートカット表示ボタン
    ' ショートカット表示フラグをセット
    ShortCutFlag = False
    Unload Me
End Sub

Private Sub CommandButton12_Click()
' マクロなしで保存終了ボタン
    Me.Hide
    ' マクロなしで保存終了処理を呼び出す
    Call SaveWOMacro
    Unload Me
End Sub

Private Sub CommandButton13_Click()
' 撮影日表示ボタン
    Me.Hide
    ' 撮影日時表示シーケンスを呼び出す
    Call PictureDateDispSequence
    Unload Me
End Sub

Private Sub CommandButton14_Click()
' 末尾のページ削除ボタン
    Me.Hide
    ' 末尾のページ削除処理を呼び出す
    Call DeleteLastPages
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期設定
    Dim PictureExistStatus As Boolean
    ' ショートカット表示フラグをリセット
    ShortCutFlag = True
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' ボタンを無効にする
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton9.Enabled = False
        CommandButton10.Enabled = False
        CommandButton12.Enabled = False
        CommandButton13.Enabled = False
        CommandButton14.Enabled = False
        ' 処理を終了
        Exit Sub
    End If
    ' 「Ａ列」のセルが結合されている場合
    If (ActiveCell.MergeArea.Rows.Count = 10) And (ActiveCell.MergeArea.Columns.Count = 1) And _
        (ActiveCell.Row <= 65463) And (ActiveCell.Column = 1) Then
        ' 現在のページ番号が「０」以下の場合
        If pageNo(ActiveCell.Row, ActiveCell.Column) <= 0 Then
            ' ここに取込ボタンを無効にする
            CommandButton1.Enabled = False
            ' 貼付ボタンを無効にする
            CommandButton5.Enabled = False
            ' 余白コマをつめるボタンを無効にする
            CommandButton7.Enabled = False
            ' 余白コマを追加ボタンを無効にする
            CommandButton8.Enabled = False
        ' 現在のページ番号が「１」以上の場合
        Else
            ' 現在のセルに写真がある場合
            PictureExistStatus = PictureExist(ActiveCell.Row, ActiveCell.Column)
            ' 余白コマをつめるボタンを無効にする
            CommandButton7.Enabled = Not PictureExistStatus
            ' クリップボードが図形の場合
            If Application.ClipboardFormats(1) = xlClipboardFormatPICT Then
                ' 現在のセルに写真がある場合は貼付ボタンを無効にする
                CommandButton5.Enabled = Not PictureExistStatus
            ' クリップボードが図形でない場合
            Else
                ' 貼付ボタンを無効にする
                CommandButton5.Enabled = False
            End If
        End If
    ' セルが結合されていない場合
    Else
        ' ここに取込ボタンを無効にする
        CommandButton1.Enabled = False
        ' 貼付ボタンを無効にする
        CommandButton5.Enabled = False
        ' 余白コマをつめるボタンを無効にする
        CommandButton7.Enabled = False
        ' 余白コマを追加ボタンを無効にする
        CommandButton8.Enabled = False
    End If
    ' 写真を切り取りした時のシート名が設定されていなければ
    If CutDataSheet = "" Then
        ' 貼り付けボタンを無効にする
        CommandButton5.Enabled = False
    End If
    ' 撮影日表示ボタン名を設定
    If Range(PictureDateFlag).Value <> 0 Then
        CommandButton13.Caption = "撮影日表示ＯＦＦ(F)"
    Else
        CommandButton13.Caption = "撮影日表示ＯＮ(F)"
    End If
End Sub
