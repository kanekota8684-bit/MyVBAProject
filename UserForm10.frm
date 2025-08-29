VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm10 
   Caption         =   "画像選択（順番にクリックしてください）"
   ClientHeight    =   10728
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18228
   OleObjectBlob   =   "UserForm10.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim clickCount As Long
Public FormClosedByX As Boolean

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.StartUpPosition = 1
    clickCount = 0
    FormClosedByX = False

    Dim i As Long, ext As String

    If Not IsArray(SelectedPaths) Then Exit Sub

    For i = 1 To 36
        If i <= UBound(SelectedPaths) Then
            If Dir(SelectedPaths(i)) = "" Then
                Me.Controls("Image" & i).Visible = False
                Me.Controls("Image" & i).Enabled = False
                Me.Controls("Label" & i).Visible = False
                SelectedOrder(i) = 0
            Else
                ext = LCase(Right(SelectedPaths(i), Len(SelectedPaths(i)) - InStrRev(SelectedPaths(i), ".")))
                If ext = "png" Or ext = "jpg" Or ext = "jpeg" Or ext = "bmp" Or ext = "gif" Then
                    Me.Controls("Image" & i).Picture = LoadPicture(SelectedPaths(i))
                    Me.Controls("Image" & i).Visible = True
                    Me.Controls("Image" & i).Enabled = True
                Else
                    Me.Controls("Image" & i).Visible = False
                    Me.Controls("Image" & i).Enabled = False
                End If
            End If
        Else
            Me.Controls("Image" & i).Visible = False
            Me.Controls("Image" & i).Enabled = False
        End If
        Me.Controls("Label" & i).Caption = ""
        Me.Controls("Label" & i).Visible = False
        SelectedOrder(i) = 0
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then FormClosedByX = True
End Sub

Private Sub CommandButtonClose_Click()
    Me.Hide
End Sub

' Imageクリックイベント（1～36）
Private Sub Image1_Click(): RegisterSelection 1: End Sub
Private Sub Image2_Click(): RegisterSelection 2: End Sub
Private Sub Image3_Click(): RegisterSelection 3: End Sub
Private Sub Image4_Click(): RegisterSelection 4: End Sub
Private Sub Image5_Click(): RegisterSelection 5: End Sub
Private Sub Image6_Click(): RegisterSelection 6: End Sub
Private Sub Image7_Click(): RegisterSelection 7: End Sub
Private Sub Image8_Click(): RegisterSelection 8: End Sub
Private Sub Image9_Click(): RegisterSelection 9: End Sub
Private Sub Image10_Click(): RegisterSelection 10: End Sub
Private Sub Image11_Click(): RegisterSelection 11: End Sub
Private Sub Image12_Click(): RegisterSelection 12: End Sub
Private Sub Image13_Click(): RegisterSelection 13: End Sub
Private Sub Image14_Click(): RegisterSelection 14: End Sub
Private Sub Image15_Click(): RegisterSelection 15: End Sub
Private Sub Image16_Click(): RegisterSelection 16: End Sub
Private Sub Image17_Click(): RegisterSelection 17: End Sub
Private Sub Image18_Click(): RegisterSelection 18: End Sub
Private Sub Image19_Click(): RegisterSelection 19: End Sub
Private Sub Image20_Click(): RegisterSelection 20: End Sub
Private Sub Image21_Click(): RegisterSelection 21: End Sub
Private Sub Image22_Click(): RegisterSelection 22: End Sub
Private Sub Image23_Click(): RegisterSelection 23: End Sub
Private Sub Image24_Click(): RegisterSelection 24: End Sub
Private Sub Image25_Click(): RegisterSelection 25: End Sub
Private Sub Image26_Click(): RegisterSelection 26: End Sub
Private Sub Image27_Click(): RegisterSelection 27: End Sub
Private Sub Image28_Click(): RegisterSelection 28: End Sub
Private Sub Image29_Click(): RegisterSelection 29: End Sub
Private Sub Image30_Click(): RegisterSelection 30: End Sub
Private Sub Image31_Click(): RegisterSelection 31: End Sub
Private Sub Image32_Click(): RegisterSelection 32: End Sub
Private Sub Image33_Click(): RegisterSelection 33: End Sub
Private Sub Image34_Click(): RegisterSelection 34: End Sub
Private Sub Image35_Click(): RegisterSelection 35: End Sub
Private Sub Image36_Click(): RegisterSelection 36: End Sub

Private Sub HandleClick(index As Long)
    Dim i As Long
    If SelectedOrder(index) = 0 Then
        clickCount = clickCount + 1
        SelectedOrder(index) = clickCount
        Me.Controls("Label" & index).Caption = clickCount & "枚目"
        Me.Controls("Label" & index).Visible = True
    Else
        For i = 1 To 36
            If SelectedOrder(i) > SelectedOrder(index) Then
                SelectedOrder(i) = SelectedOrder(i) - 1
                Me.Controls("Label" & i).Caption = SelectedOrder(i) & "枚目"
            End If
        Next i
        SelectedOrder(index) = 0
        Me.Controls("Label" & index).Visible = False
        clickCount = clickCount - 1
    End If
End Sub

' 選択順を記録する共通処理
Private Sub RegisterSelection(imgIndex As Integer)
    Dim nextOrder As Integer
    nextOrder = 1

    ' 次の空き順番を探す
    Do While nextOrder <= 36
        If Not IsInArray(nextOrder, SelectedOrder) Then Exit Do
        nextOrder = nextOrder + 1
    Loop

    ' 登録
    SelectedOrder(imgIndex) = nextOrder
    Me.Controls("Label" & imgIndex).Caption = Format(nextOrder) & "枚目"
End Sub

' 配列に値が含まれているか判定
Private Function IsInArray(val As Variant, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Private Sub CommandButton1_Click()
    Dim i As Integer, j As Integer
    Dim imgPath As String
    Dim pasteCell As Range
    Dim targetCell As Range
    Dim mergedCell As Range
    Dim orderIndex As Integer

    If PasteStartCell Is Nothing Then Exit Sub
    Set pasteCell = PasteStartCell

    ' 選択順に並び替えて貼り付け
    For orderIndex = 1 To UBound(SelectedOrder)
    For i = 1 To UBound(SelectedOrder)
        If SelectedOrder(i) = orderIndex Then
            imgPath = SelectedPaths(i)
            Set targetCell = pasteCell.Offset(orderIndex - 1, 0)
            Set mergedCell = targetCell.MergeArea



                ' 画像貼り付け
                Dim pic As Shape
                Set pic = ActiveSheet.Shapes.AddPicture(Filename:=imgPath, _
                    LinkToFile:=False, SaveWithDocument:=True, _
                    Left:=mergedCell.Left, Top:=mergedCell.Top, _
                    Width:=0, Height:=0)

                With pic
                    .ZOrder msoSendToBack
                    .LockAspectRatio = msoTrue
                    .ScaleHeight 1, msoTrue
                    .ScaleWidth 1, msoTrue

                    ' サイズ調整
                    Dim scaleRatio As Double
                    scaleRatio = Application.Min(mergedCell.Width / .Width, mergedCell.Height / .Height)
                    .ScaleWidth scaleRatio, msoTrue
                    .ScaleHeight scaleRatio, msoTrue

                    ' 中央配置
                    .Left = mergedCell.Left + (mergedCell.Width - .Width) / 2
                    .Top = mergedCell.Top + (mergedCell.Height - .Height) / 2
                End With
                Exit For
            End If
        Next i
    Next orderIndex

    Unload Me
End Sub

