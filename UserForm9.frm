VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "撮影日の編集"
   ClientHeight    =   1090
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3876
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myPictureNo As Integer, myPictureName As String

Private Sub CommandButton1_Click()
' 決定ボタン
    Dim myDate As String, myType As Integer
    myDate = Format(TextBox1.Value) & "/" & Format(TextBox2.Value) & "/" & Format(TextBox3.Value) _
        & " " & Format(TextBox4.Value) & ":" & Format(TextBox5.Value)
    ' 値のチェック
    If Not IsDate(myDate) Then
        MsgBox "値が不正です。", vbOKOnly + vbExclamation, "お知らせ"
        Exit Sub
    End If
    ' 値を格納
    Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value = myDate
    ' 撮影日の表示形式フラグをエンコード
    myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
    ' テキストボックスの文字を書き換え
    With ActiveSheet.Shapes(Application.Caller)
        .TextFrame.Characters.Text = PictureDateFormat(myDate, myType)
    End With
    ' テキストボックスの位置決め
    Call PictureDatePosition(myPictureName)
    ' ユーザーフォームを解放
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' キャンセルボタン
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期化処理
    Dim myPictureDate As String
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' 各コントロールを無効にする
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        CommandButton1.Enabled = False
        Exit Sub
    End If
    ' 写真のページ番号を求める
    myPictureNo = GetPictureNoFromDate(Application.Caller)
    ' 写真の撮影日データを取得
    myPictureDate = PictureDateFormat(Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value, 0)
    ' 写真の名前を取得
    myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value
    ' 撮影日データが日付データなら
    If IsDate(myPictureDate) Then
        ' テキストボックスの初期値を設定
        TextBox1.Value = Format(myPictureDate, "yyyy")
        TextBox2.Value = Format(myPictureDate, "m")
        TextBox3.Value = Format(myPictureDate, "d")
        TextBox4.Value = Format(myPictureDate, "h")
        TextBox5.Value = Format(myPictureDate, "n")
    ' 日付データとして扱えない場合
    Else
        TextBox1.Value = "****"
        TextBox2.Value = "**"
        TextBox3.Value = "**"
        TextBox4.Value = "**"
        TextBox5.Value = "**"
    End If
End Sub
