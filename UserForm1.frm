VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "印刷する前の処理を選択してください。"
   ClientHeight    =   1320
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' 印刷ボタン
    Me.Hide
    Call PreviewPrint
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' キャンセルボタン
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' コマ毎に番号振直しボタン
    Me.Hide
    Call SerialNumbering
    Call PreviewPrint
    Unload Me
End Sub

Private Sub CommandButton4_Click()
' 写真毎に番号振直しボタン
    Me.Hide
    Call PictureNumbering
    Call PreviewPrint
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期設定
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' ボタンを無効にする
        CommandButton3.Enabled = False
        CommandButton4.Enabled = False
    End If
End Sub
