VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "処理を選択してください。"
   ClientHeight    =   1130
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4128
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' ここに写真を移動ボタン
    Call MoveToHere
    ' 写真の移動元と移動先ページ番号をクリア
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' この写真と入れ替えボタン
    Call ExchangePicture
    ' 写真の移動元と移動先ページ番号をクリア
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' キャンセルボタン
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' 写真の移動元ページのセルを選択
    Worksheets(SwapSourceSheet).Select
    Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
    ' ユーザーフォーム５をモードレスで表示
    UserForm5.Show vbModeless
    ' ユーザーフォームを解放
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期処理
    ' 左クリックされた写真のセルを選択
    ActiveSheet.Shapes(Application.Caller).TopLeftCell.MergeArea.Select
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' ボタンを無効にする
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ユーザーフォームの終了前の処理
    ' 「×」で閉じようとした場合
    If CloseMode = vbFormControlMenu Then
        ' ユーザーフォームを非表示にする
        Me.Hide
        ' 写真の移動元ページのセルを選択
        Worksheets(SwapSourceSheet).Select
        Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
        ' ユーザーフォーム５をモードレスで表示
        UserForm5.Show vbModeless
    End If
End Sub
