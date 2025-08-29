VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "写真の入れ替え処理中です。他の操作をしないでください。"
   ClientHeight    =   1280
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5400
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' キャンセルボタン
    ' 写真の移動元ページのセルを選択
    Worksheets(SwapSourceSheet).Select
    Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
    ' 写真の移動元と移動先ページ番号をクリア
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' 写真を末尾へ移動ボタン
    ' 現在のシート名を取得
    SwapDestSheet = ActiveSheet.Name
    ' 写真を末尾に移動させる
    Call MoveToEnd
    ' 写真の移動元と移動先ページ番号をクリア
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期処理
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' ボタンを無効にする
        CommandButton2.Enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ユーザーフォームの終了前の処理
    ' 「×」で閉じようとした場合
    If CloseMode = vbFormControlMenu Then
        ' 写真の移動元ページのセルを選択
        Worksheets(SwapSourceSheet).Select
        Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
        ' 写真の移動元と移動先ページ番号をクリア
        SwapSourceNo = 0
        SwapDestNo = 0
    End If
End Sub
