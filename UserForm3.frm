VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "処理を選択してください。"
   ClientHeight    =   2540
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' 左回転ボタン
    ' 回転処理（図の名前、回転角度）を呼び出す
    Call RotationProc(Application.Caller, -90)
    ' 撮影日テキストボックスの位置決め
    If Range(PictureDateFlag).Value <> 0 Then
        Call PictureDatePosition(Application.Caller)
    End If
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' 右回転ボタン
    ' 回転処理（図の名前、回転角度）を呼び出す
    Call RotationProc(Application.Caller, 90)
    ' 撮影日テキストボックスの位置決め
    If Range(PictureDateFlag).Value <> 0 Then
        Call PictureDatePosition(Application.Caller)
    End If
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' キャンセルボタン
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton4_Click()
' 切り取りボタン
    Dim CurrentNo As Integer, CurrentName As String
    On Error Resume Next
    CurrentName = ActiveSheet.Shapes(Application.Caller).Name
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag) <> 0 Then
        ' 撮影日テキストボックスを削除
        ActiveSheet.Shapes(PictureDateName(CurrentName)).Delete
    End If
    ' 図を切り取る
    ActiveSheet.Shapes(Application.Caller).Cut
    ' 撮影日データを切り取りバッファへ保存
    Range(PictureNameBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value
    Range(PictureDateBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value
    ' 撮影日データを消去
    Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    ' 現在のシート名を記憶
    CutDataSheet = ActiveSheet.Name
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton5_Click()
' ここに余白コマを追加ボタン
    Me.Hide
    ' コマの追加を呼び出す
    Call AddBlank
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton6_Click()
' 写真を削除して詰めるボタン
    Dim CurrentNo As Integer, CurrentName As String
    On Error Resume Next
    CurrentName = ActiveSheet.Shapes(Application.Caller).Name
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    Me.Hide
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag) <> 0 Then
        ' 撮影日テキストボックスを削除
        ActiveSheet.Shapes(PictureDateName(CurrentName)).Delete
    End If
    ' 図を切り取る
    ActiveSheet.Shapes(Application.Caller).Cut
    ' 撮影日データを切り取りバッファに移す
    Range(PictureNameBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value
    Range(PictureDateBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value
    ' 撮影日データを消去
    Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    ' 現在のシートを記憶
    CutDataSheet = ActiveSheet.Name
    ' コマの削除を呼び出す
    Call DeleteBlank
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton7_Click()
' ここに取込ボタン
    Me.Hide
    ' ここに一括取込を呼び出す
    Call GetMultiPictureFromHere
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton8_Click()
' ページ追加ボタン
    Me.Hide
    ' ページ追加を呼び出す（引数は追加枚数）
    Call AddPageProc(1)
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton9_Click()
' 印刷ボタン
    Me.Hide
    ' 印刷プレビューを呼び出す
    UserForm1.Show vbModal
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton10_Click()
' コマ毎に番号振直しボタン
    Me.Hide
    ' コマ毎の番号付番処理を呼び出す
    Call SerialNumbering
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton11_Click()
' 写真毎に番号振直しボタン
    Me.Hide
    ' 写真毎の番号付番処理を呼び出す
    Call PictureNumbering
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton12_Click()
' 続きに取込ボタン
    Me.Hide
    ' 一括取込を呼び出す
    Call GetMultiPicture
    ' ユーザーフォームを解放する
    Unload Me
End Sub

Private Sub CommandButton13_Click()
' この写真を入れ替えボタン
    ' 写真の移動元のページ番号と写真の名前と現在のシート名を設定
    SwapSourceNo = pageNo(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row, ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column)
    SwapSourceName = ActiveSheet.Shapes(Application.Caller).Name
    SwapSourceSheet = ActiveSheet.Name
    ' ユーザーフォームを非表示にする
    Me.Hide
    ' ユーザーフォーム５をモードレスで表示
    UserForm5.Show vbModeless
    ' ユーザーフォームの解放
    Unload Me
End Sub

Private Sub CommandButton14_Click()
' 撮影日表示ボタン
    Me.Hide
    ' 撮影日時表示シーケンスを呼び出す
    Call PictureDateDispSequence
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ユーザーフォームの初期設定
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' ボタンを無効にする
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        CommandButton4.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton10.Enabled = False
        CommandButton11.Enabled = False
        CommandButton12.Enabled = False
        CommandButton13.Enabled = False
        CommandButton14.Enabled = False
        ' 処理を終了
        Exit Sub
    End If
    With ActiveSheet.Shapes(Application.Caller)
        ' 写真の貼り付けられているセルが１０個結合されていて写真の角度が９０°刻みの場合
        If (.TopLeftCell.MergeArea.Count = 10) And (.TopLeftCell.MergeArea.Columns.Count = 1) And _
            ((.Rotation = 0) Or (.Rotation = 90) Or (.Rotation = 180) Or (.Rotation = 270)) Then
            ' 写真のセルを選択する
            .TopLeftCell.MergeArea.Select
        ' 写真の貼り付けられているセルが１０個結合されていないまたは写真の角度が９０°刻みでない場合
        Else
            ' ボタンを無効にする
            CommandButton1.Enabled = False
            CommandButton2.Enabled = False
            CommandButton5.Enabled = False
            CommandButton6.Enabled = False
            CommandButton7.Enabled = False
            CommandButton8.Enabled = False
            CommandButton9.Enabled = False
            CommandButton10.Enabled = False
            CommandButton11.Enabled = False
            CommandButton12.Enabled = False
            CommandButton13.Enabled = False
            CommandButton14.Enabled = False
        End If
    End With
    ' 撮影日表示ボタン名を設定
    If Range(PictureDateFlag).Value <> 0 Then
        CommandButton14.Caption = "撮影日表示ＯＦＦ(F)"
    Else
        CommandButton14.Caption = "撮影日表示ＯＮ(F)"
    End If
End Sub
