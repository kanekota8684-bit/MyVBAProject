VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
' ワークシートを開いたときの処理
    ' 写真の移動元と移動先ページ番号をクリア
    SwapSourceNo = 0
    SwapDestNo = 0
End Sub
