VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
' ダブルクリックした場合の処理
    ' 現在のシートを選択
    ActiveSheet.Select
    ' ダブルクリックされたセルを選択
    Target.Select
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' 処理を終了
        Exit Sub
    End If
    ' 「Ａ列」のセルが結合されていれば
    If (ActiveCell.MergeArea.Rows.Count = 10) And _
        (ActiveCell.MergeArea.Columns.Count = 1) And _
        (ActiveCell.Row <= 65463) And (ActiveCell.Column = 1) And _
        SwapSourceNo = 0 Then
        ' ここに一括取込を呼び出す
        Call GetMultiPictureFromHere
        ' 入力モードになるのをキャンセル
        Cancel = True
    ' 写真の入れ替え選択中なら
    ElseIf SwapSourceNo <> 0 Then
        ' 入力モードになるのをキャンセル
        Cancel = True
    End If
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
' 右クリックした場合の処理
    ' 現在のシートを選択
    ActiveSheet.Select
    ' 右クリックされたセルを選択
    Target.Select
    ' 「Ｂ列」のコメントエリアで右クリックされた場合のみ
    If (ActiveCell.Column = 2) And (ActiveCell.MergeCells = False) And ((ActiveCell.Row Mod 11) > 3) And _
        SwapSourceNo = 0 Then
        ' ユーザーフォーム４を呼び出す
        UserForm4.Show vbModal
    ' それ以外のセルで右クリックされた場合
    ElseIf SwapSourceNo = 0 Then
        ' ユーザーフォーム２を呼び出す
        UserForm2.Show vbModal
    ' 写真の入れ替え選択中なら
    ElseIf SwapSourceNo <> 0 Then
        ' ショートカットメニューを表示しない
        ShortCutFlag = True
    End If
    ' ショートカットメニューの表示を制御
    Cancel = ShortCutFlag
End Sub
