Attribute VB_Name = "Module1"
Option Explicit
Const myMinRow As Long = 2
Const MaxPageRow As Long = 65472
Const MaxPageNo As Integer = 5952
Const IndexRowHeight As Double = 17.25 ' 番号行のセルの高さ
Const CommentRowHeight As Double = 30.75 ' 行のセルの高さ
Public Const TopRowHeight As Double = 15# ' １行目のセルの高さ
Public Const TempRowHeight As Double = 300# ' １行目で画像がズレる対策用セルの高さ
Public ShortCutFlag As Boolean ' ショートカット表示用フラグ
Public SwapSourceNo As Integer, SwapDestNo As Integer ' 写真の入れ替え元、入れ替え先のページ番号
Public SwapSourceName As String, SwapDestName As String ' 写真の入れ替え元、入れ替え先の写真名
Public SwapSourceSheet As String, SwapDestSheet As String ' 写真の入れ替え元、入れ替え先のシート名
Public ComboList1 As Long ' コンボボックス１のリスト番号保存用（ユーザーフォーム４で使用）
Public ComboList2 As Long ' コンボボックス２のリスト番号保存用（ユーザーフォーム４で使用）
Public ComboList3 As Long ' コンボボックス３のリスト番号保存用（ユーザーフォーム４で使用）
Public ComboList4 As Long ' コンボボックス４のリスト番号保存用（ユーザーフォーム４で使用）
Public ComboList5 As Long ' コンボボックス５のリスト番号保存用（ユーザーフォーム４で使用）
Public ComboList6 As Long ' コンボボックス６のリスト番号保存用（ユーザーフォーム４で使用）
Public Const PictureNameBuffer As String = "K" ' 写真のファイル名保存用セルの列番号
Public Const PictureDateBuffer As String = "L" ' 写真の撮影日保存用セルの列番号
Public Const MinDataBuffer As Integer = 12 ' 写真の撮影日、ファイル名保存用セルの開始行番号オフセット
Public Const CutDataBuffer As Integer = 12 ' 切り取った写真の撮影日、ファイル名保存用セルの行番号
Public Const PictureDateFlag As String = "K2" ' 写真の撮影日の表示をするかどうかのフラグ保存用セル番地
Public Const PictureDateType As String = "L2" ' 撮影日表示タイプのフラグ保存用セル番地
Public Const DateSuffix As String = "D" ' 撮影日テキストボックスの名前識別用文字
Public CutDataSheet As String ' 写真の切り取りを行ったシート名
Public Const DateFontSize As String = "K3" ' 撮影日のフォントサイズ保存用のセル番地
Public Const DateFontBold As String = "K4" ' 撮影日のフォント太字フラグ保存用のセル番地
Public Const DateFontColorR As String = "L3" ' 撮影日のフォント色（赤）保存用のセル番地
Public Const DateFontColorG As String = "L4" ' 撮影日のフォント色（緑）保存用のセル番地
Public Const DateFontColorB As String = "L5" ' 撮影日のフォント色（青）保存用のセル番地
Public Const DateHeightOffset As Double = 0# ' 撮影日テキストボックスの高さマージン
Public Const DateXOffset As String = "K5" ' 撮影日テキストボックスの右からのオフセット保存用セル番地
Public Const DateYOffset As String = "K6" ' 撮影日テキストボックスの下からのオフセット保存用セル番地
Public Const XUnit As Double = 0.33 ' 右からのオフセットの単位変換値
Public Const YUnit As Double = 0.325 ' 下からのオフセットの単位変換値
Public Const DateSeparator As String = "L6" ' 撮影日の区切りを「．」にするフラグ保存用のセル番地
Public Const WeekDisp As String = "K7" ' 撮影日に曜日を表示するフラグ保存用のセル番地
Public Const WeekLang As String = "K8" ' 撮影日の曜日の言語フラグ保存用のセル番地
Public SelectedOrder() As Integer ' 画像の順番（例：Image1が2枚目なら SelectedOrder(1) = 2）
Public SelectedPaths() As String  ' 選択された画像のパス
Public PasteStartCell As Range ' ← 標準モジュールに宣言

Function AddPages(myInsertCount As Integer) As Integer
' 原紙を追加する関数（引数は追加する枚数）
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim myPageCount As Integer
    Dim myRange As Range
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' 戻り値を－１にする
        AddPages = -1
        ' 処理を終了
        Exit Function
    End If
    ' ページカウンタ
    j = 1
    ' ページ挿入枚数
    k = 0
    ' 最終ページ
    m = 1
    ' 先頭のページのセルを選択する
    Set myRange = Range(PictureColumn(1) & Format(PictureRow(1)))
    myRange.Select
    ' ステータスバーに状態を表示する
    Application.StatusBar = "ページを追加しています。お待ちください。"
    ' 繰り返し処理
    Do
        ' セルが１０個結合されている場合
        If (myRange.MergeArea.Rows.Count = 10) And (myRange.MergeArea.Columns.Count = 1) Then
            ' ページカウンタを加算
            j = j + 1
            ' 最終ページを記憶
            m = j
            ' 次のページのセルを選択
            Set myRange = Range(PictureColumn(j) & Format(PictureRow(j)))
            myRange.Select
            ' 追加するページが最大値を超えた場合
            If j > MaxPageNo Then
                ' 戻り値を－２にする
                AddPages = -2
                Range(PictureColumn(MaxPageNo) & Format(PictureRow(MaxPageNo))).Select
                Application.StatusBar = False
                ' 処理を終了
                Exit Function
            End If
        ' セルが結合されていない場合または既定の結合でない場合
        Else
            ' セルが結合されている場合
            If myRange.MergeCells Then
                ' セルの結合を解除
                myRange.MergeCells = False
            End If
            ' 行の高さを設定する
            myRange.Offset(-1, 0).RowHeight = TopRowHeight
            ' 追加する行数を計算する
            myPageCount = (3 - ((j - 1) Mod 3))
            ' 原紙を追加
            For i = 1 To myPageCount
                ' セルを結合
                Range(PictureColumn(j) & Format(PictureRow(j)) & _
                    ":" & PictureColumn(j) & Format(PictureRow(j) + 9)).Merge
                ' 余白文字を挿入する
                With Range(PictureColumn(j) & Format(PictureRow(j))).MergeArea
                    ' セルの値
                    .Value = "余白"
                    ' フォントサイズ
                    .Font.Size = 72
                    ' 文字の色
                    .Font.Color = RGB(192, 192, 192)
                    ' 水平位置を中央合わせ
                    .HorizontalAlignment = xlCenter
                    ' 垂直位置を中央合わせ
                    .VerticalAlignment = xlCenter
                End With
                ' 行の高さを設定する
                Range("A" & Format(PictureRow(j)) & ":A" & Format(PictureRow(j) + 1)).RowHeight = IndexRowHeight
                ' 罫線を引く
                With Range(CommentColumn(j) & Format(PictureRow(j)))
                    ' 直線を引く
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    ' 罫線の色
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                    ' ページ番号を挿入する
                    .Value = "No." & Format(j)
                    ' 水平位置を左詰め
                    .HorizontalAlignment = xlLeft
                    ' 垂直位置を中央合わせ
                    .VerticalAlignment = xlCenter
                    ' フォントサイズ
                    .Font.Size = 11
                    ' 文字の色
                    .Font.Color = RGB(0, 0, 0)
                End With
                With Range(CommentColumn(j) & Format(PictureRow(j) + 2) & _
                    ":" & CommentColumn(j) & Format(PictureRow(j) + 8))
                    ' 選択範囲の下に点線を引く
                    .Borders(xlEdgeBottom).LineStyle = xlDot
                    ' 罫線の色
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                    ' 選択範囲の中に点線を引く
                    .Borders(xlInsideHorizontal).LineStyle = xlDot
                    ' 罫線の色
                    .Borders(xlInsideHorizontal).Color = RGB(0, 0, 0)
                    ' 水平位置
                    .HorizontalAlignment = xlLeft
                    ' 垂直位置
                    .VerticalAlignment = xlCenter
                    ' フォントサイズ
                    .Font.Size = 11
                    ' 文字の色
                    .Font.Color = RGB(0, 0, 0)
                    ' 縮小して全体を表示
                    .WrapText = False
                    .ShrinkToFit = True
                    ' 行の高さを設定する
                    .RowHeight = CommentRowHeight
                End With
                ' 行の高さを設定する
                Range("A" & Format(PictureRow(j) + 9)).RowHeight = CommentRowHeight
                Range("A" & Format(PictureRow(j) + 10)).RowHeight = TopRowHeight
                ' 次のページへ
                j = j + 1
                ' 次のページのセルを選択
                Set myRange = Range(PictureColumn(j) & Format(PictureRow(j)))
                myRange.Select
            Next i
            ' ページ挿入枚数を加算
            k = k + 1
            ' ページ挿入枚数が引数と一致した場合
            If k >= myInsertCount Then
                ' 繰り返し処理を終わる
                Exit Do
            End If
            ' 追加するページが最大値を超えた場合
            If j > MaxPageNo Then
                ' 戻り値を－２にする
                AddPages = -2
                Range(PictureColumn(MaxPageNo) & Format(PictureRow(MaxPageNo))).Select
                Application.StatusBar = False
                ' 処理を終了
                Exit Function
            End If
        End If
    Loop
    ' 追加した原紙の先頭のセルを選択する
    Range(PictureColumn(m) & Format(PictureRow(m))).MergeArea.Select
    ' 戻り値をゼロにする（正常終了）
    AddPages = 0
    Application.StatusBar = False
End Function

Sub PictureRotation()
' 写真を左クリックした場合の処理
    ' 現在のシートを選択
    ActiveSheet.Select
    ' 写真の移動元ページ番号がセットされていなければ
    If SwapSourceNo = 0 Then
        ' ユーザーフォーム３を呼び出す
        UserForm3.Show vbModal
    ' 写真の移動元ページ番号がセットされていれば
    Else
        ' 写真の入れ替え先ページ番号を取得
        SwapDestNo = pageNo(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row, ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column)
        ' 写真の入れ替え先写真名を取得
        SwapDestName = ActiveSheet.Shapes(Application.Caller).Name
        ' 写真の入れ替え先シート名を取得
        SwapDestSheet = ActiveSheet.Name
        ' 写真の入れ替え元ページ番号と写真の入れ替え先ページ番号が異なれば
        If (SwapSourceNo <> SwapDestNo) Or (SwapSourceSheet <> SwapDestSheet) Then
            ' ユーザーフォーム５を非表示にし
            UserForm5.Hide
            ' ユーサーフォーム６を呼び出す
            UserForm6.Show vbModal
        ' 写真の入れ替え元ページ番号と写真の入れ替え先ページ番号が同じならば
        Else
            ' メッセージを表示
            MsgBox "写真の入れ替え元と写真の入れ替え先が同じです。", vbOKOnly + vbExclamation, "お知らせ"
        End If
    End If
End Sub

Sub PastePicture(ByVal myFilename As String)
    Dim myPictureName As String, myPictureNo As Integer, myDate As String, myType As Integer
    Dim cellWidth As Double, cellHeight As Double
    Dim scaleRatio As Double

    myPictureName = Right(myFilename, Len(myFilename) - InStrRev(myFilename, "\"))

    ActiveSheet.Shapes. _
        AddPicture(Filename:=myFilename, _
                    LinkToFile:=False, SaveWithDocument:=True, _
                    Left:=ActiveCell.MergeArea.Left, Top:=ActiveCell.MergeArea.Top, _
                    Width:=0, Height:=0).Select

    With Selection.ShapeRange
        .ZOrder msoSendToBack
        .Name = myPictureName
        .LockAspectRatio = msoTrue

        ' セルサイズ取得
        cellWidth = ActiveCell.MergeArea.Width - 4
        cellHeight = ActiveCell.MergeArea.Height - 4

        ' 元サイズに戻す
        .ScaleHeight 1, msoTrue
        .ScaleWidth 1, msoTrue

        ' セルに収まるように縦横比を維持して縮小
        If .Width > cellWidth Or .Height > cellHeight Then
            scaleRatio = Application.Min(cellWidth / .Width, cellHeight / .Height)
            .ScaleWidth scaleRatio, msoTrue
            .ScaleHeight scaleRatio, msoTrue
        End If

        ' 中央に配置
        .Left = ActiveCell.MergeArea.Left + (ActiveCell.MergeArea.Width - .Width) / 2
        .Top = ActiveCell.MergeArea.Top + (ActiveCell.MergeArea.Height - .Height) / 2
    End With

    Selection.OnAction = "PictureRotation"
    myPictureNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    myDate = PictureDate(myFilename)

    Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value = myPictureName
    Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value = myDate

    If Range(PictureDateFlag).Value <> 0 Then
        ActiveCell.MergeArea.Select
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If

    ActiveCell.MergeArea.Select
End Sub
Sub RotationProc(ByVal myShapeName As String, ByVal myDegree As Integer)
' 写真の回転処理（引数は図形の名前、回転角度）
    Dim myWidth As Double, myHeight As Double, myAspectRatio As Double
    Application.ScreenUpdating = False
    With ActiveSheet.Shapes(myShapeName)
        .TopLeftCell.MergeArea.Select
        ' 「Ａ２」セルで写真が縦長で９０°または２７０°の場合に位置決めがずれる対策
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TempRowHeight
        End If
        myWidth = .Width
        myHeight = .Height
        ' 写真の縦横比を計算
        myAspectRatio = myWidth / myHeight
        ' 縦横比の固定を解除する
        .LockAspectRatio = msoFalse
        ' 写真を正方形にする
        If myWidth > myHeight Then
            .Height = myWidth
        ElseIf myWidth < myHeight Then
            .Width = myHeight
        End If
        ' 写真を回転させる
        .Rotation = .Rotation + myDegree
        ' 縦横比を復元する
        If .Rotation = 90 Or .Rotation = 270 Then
        ' 回転角度が９０°または２７０°の場合
            If myWidth > myHeight Then
            ' 写真の幅が高さより大きい場合
                ' 写真の幅を枠に合わせる
                .Width = Int(Selection.Height) - 4
                ' 写真の高さを復元
                .Height = .Width / myAspectRatio
            ElseIf myWidth < myHeight Then
            ' 写真の高さが幅より大きい場合
                ' 写真の高さを枠に合わせる
                .Height = Int(Selection.Width) - 4
                ' 写真の幅を復元
                .Width = .Height * myAspectRatio
            End If
            ' 縦横比を固定する
            .LockAspectRatio = msoTrue
            ' 写真の幅または高さがセルより大きくなればセルに合わせる
            If .Width > Int(Selection.Height) - 4 Then
                .Width = Int(Selection.Height) - 4
            ElseIf .Height > Int(Selection.Width) - 4 Then
                .Height = Int(Selection.Width) - 4
            End If
        Else
        ' 回転角度が０°または１８０°の場合
            If myWidth > myHeight Then
            ' 写真の幅が高さより大きい場合
                ' 写真の幅を枠に合わせる
                .Width = Int(Selection.Width) - 4
                ' 写真の高さを復元
                .Height = .Width / myAspectRatio
            ElseIf myWidth < myHeight Then
            ' 写真の高さが幅より大きい場合
                ' 写真の高さを枠に合わせる
                .Height = Int(Selection.Height) - 4
                ' 写真の幅を復元
                .Width = .Height * myAspectRatio
            End If
            ' 縦横比を固定する
            .LockAspectRatio = msoTrue
            ' 写真の幅または高さがセルより大きくなればセルに合わせる
            If .Height > Int(Selection.Height) - 4 Then
                .Height = Int(Selection.Height) - 4
            ElseIf .Width > Int(Selection.Width) - 4 Then
                .Width = Int(Selection.Width) - 4
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
    Application.ScreenUpdating = True
End Sub

Sub PreviewPrint()
' 印刷プレビューを表示する
    Dim myTopCount As Long, myBottomCount As Long
    Dim myPictureNo As Integer, myMaxNo As Integer
    Dim myPicture As Shape
    ' 写真の最大ページ番号
    myMaxNo = 1
    ' すべての図にたいして
    For Each myPicture In ActiveSheet.Shapes
        myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' 図の行数がページの最大値より大きい場合
        If myPicture.BottomRightCell.Row > MaxPageRow Then
            ' メッセージを表示する
            MsgBox "写真が最大ページを超えて貼り付けられています。", vbOKOnly + vbExclamation, "お知らせ"
            ' 処理を終了
            Exit Sub
        ElseIf (myPictureNo > myMaxNo) Then
            ' 図の最大行数を求める
            myMaxNo = myPictureNo
        End If
    Next
    myTopCount = 1
    myBottomCount = (PictureRow(myMaxNo) \ 33) * 33 + 33
    With ActiveSheet
        ' 印刷範囲を設定する
        .PageSetup.PrintArea = "A" & Format(myTopCount) & ":B" & Format(myBottomCount)
        ' 印刷の方向を縦に設定する
        .PageSetup.Order = xlDownThenOver
        ' 印刷プレビューを表示する
        .PrintPreview
        ' 印刷範囲を解除する
        .PageSetup.PrintArea = False
    End With
End Sub

Sub AddPageProc(ByVal myPage As Integer)
' ページの追加処理（引数は追加枚数）
    ' 画面表示の更新をしないようにする
    Application.ScreenUpdating = False
    ' 原紙の追加関数（追加枚数）
    Select Case AddPages(myPage)
        ' 戻り値による処理
        Case -1
            MsgBox "シートが保護されています。" _
                & vbCrLf & "保護を解除してください。", vbOKOnly + vbExclamation, "お知らせ"
        Case -2
            MsgBox "これ以上ページを追加できません。", vbOKOnly + vbExclamation, "お知らせ"
    End Select
    ' 画面表示の更新を許可する
    Application.ScreenUpdating = True
    ActiveCell.Select
End Sub

Sub GetMultiPicture()
' 一括取り込み処理
    Dim myPicture As Shape
    Dim myRange As Range
    Dim myPictureNo As Integer, myMaxNo As Integer
    ' 図形の最大ページ数
    myMaxNo = 0
    ' すべての図形に対して
    For Each myPicture In ActiveSheet.Shapes
        ' 図形の左上セルが最小セル以上の場合
        If myPicture.TopLeftCell.Row >= myMinRow Then
            ' 図形のページ数を求める
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' ページの最大数を求める
            If myMaxNo < myPictureNo Then
                myMaxNo = myPictureNo
            End If
        End If
    Next
    myMaxNo = myMaxNo + 1
    ' 最大ページを超える場合は処理を中断
    If myMaxNo > MaxPageNo Then
        MsgBox "写真が最大ページを超えます。" _
            & vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
    Else
        ' 写真のページからセルを設定
        Set myRange = Range(PictureColumn(myMaxNo) & Format(PictureRow(myMaxNo)))
        ' セルを選択
        myRange.MergeArea.Select
        ' ここに一括取込を呼び出す
        Call GetMultiPictureFromHere
    End If
End Sub

Sub GetMultiPictureFromHere()
' ここに一括取り込み処理
    Dim myFilename() As String, mySelectedItemsCount As Long
    Dim myPicture As Shape, myPictureName As String, TempName As String, PictureName() As String
    Dim myPictureNo As Integer, CurrentNo As Integer, StartNo As Integer
    Dim MinNo As Integer, MaxNo As Integer, NextNo As Integer, BlankCount As Integer
    Dim myRange As Range
    Dim CurrentRow As Long, CurrentColumn As String, NextRow As Long, NextColumn As String
    Dim i As Long, j As Long, myFailureCount As Long, myFname() As String, myFnameCount As Long
    Dim m As Integer, PageInsertCount As Integer, myLongNameCount As Long
    Dim mySamePictureExist As Boolean
    Dim myCurrentWindowZoom As Double
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        MsgBox "シートが保護されています。" _
            & vbCrLf & "保護を解除してください。", vbOKOnly + vbExclamation, "お知らせ"
        ' 処理を終了
        Exit Sub
    End If
    
    Dim targetCell As Range
    Set targetCell = ActiveCell ' 例：現在選択されているセルを対象にする
    Set PasteStartCell = targetCell
    
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    StartNo = CurrentNo
    ' 現在のページ番号が「０」以下の場合処理を終了
    If CurrentNo <= 0 Then
        ' 処理を終了
        Exit Sub
    End If
    ' ファイルダイアログボックスを開く
    
    
    
   With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = True
    .Title = "画像を選択してください"
    .ButtonName = "取り込み"
    .Filters.Clear
    .Filters.Add "画像", "*.JPG;*.JPEG;*.BMP;*.TIF;*.TIFF;*.PNG;*.GIF;*.HEIC", 1

    If .Show = -1 Then
        mySelectedItemsCount = .SelectedItems.Count

        ReDim SelectedPaths(1 To mySelectedItemsCount)
        ReDim SelectedOrder(1 To mySelectedItemsCount)
        
        ' ファイルパスを SelectedPaths に格納
    For i = 1 To mySelectedItemsCount
        SelectedPaths(i) = .SelectedItems(i)
    Next i



            ' ?? ここに初期化処理を追加！
    For i = 1 To 36
        With UserForm10.Controls("Image" & i)
            Set .Picture = Nothing
        End With
        UserForm10.Controls("Label" & i).Caption = ""
    Next i

    ' 画像表示処理（Image1～ImageN）
    Dim img As StdPicture
    Dim ctrlName As String
    For i = 1 To mySelectedItemsCount
        Set img = LoadPicture(SelectedPaths(i))
        ctrlName = "Image" & i
        With UserForm10.Controls(ctrlName)
            .Picture = img
            .PictureSizeMode = fmPictureSizeModeZoom
        End With
    Next i

    ' フォーム表示
    UserForm10.Show
End If


End With

    Application.ScreenUpdating = False
    ' 同じ名前の写真と長いファイル名の写真をカウントする
    myFailureCount = 0
    myLongNameCount = 0
    For i = 1 To UBound(myFilename)
        myPictureName = Right(myFilename(i), Len(myFilename(i)) - InStrRev(myFilename(i), "\"))
        For Each myPicture In ActiveSheet.Shapes
            If myPicture.Name = myPictureName Then
                myFailureCount = myFailureCount + 1
                Exit For
            End If
        Next
        If Len(myPictureName) > 30 Then
            myLongNameCount = myLongNameCount + 1
        End If
    Next i
    ' 取り込み枚数を計算
    myFnameCount = mySelectedItemsCount - myFailureCount - myLongNameCount
    ' 取り込み枚数が１枚以上なら
    If myFnameCount > 0 Then
        ReDim myFname(myFnameCount - 1)
        j = 0
        ' 選択ファイル数分くりかえし
        For i = 0 To mySelectedItemsCount - 1
            mySamePictureExist = False
            myPictureName = Right(myFilename(i), Len(myFilename(i)) - InStrRev(myFilename(i), "\"))
            ' 同じ名前の写真を確認
            For Each myPicture In ActiveSheet.Shapes
                If myPicture.Name = myPictureName Then
                    mySamePictureExist = True
                    Exit For
                End If
            Next
            ' 写真のファイル名が３０字以下で写真が同じ名前でなければ
            If mySamePictureExist = False And Len(myPictureName) <= 30 Then
                ' 取り込み写真を配列に代入
                myFname(j) = myFilename(i)
                j = j + 1
            End If
        Next i
        ' 写真の最大ページと枚数を求める
        MaxNo = 0
        m = 0
        ' すべての図にたいして
        For Each myPicture In ActiveSheet.Shapes
            ' 図のページを求める
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' 現在のセルより図のページが大きい場合
            If CurrentNo <= myPictureNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' 枚数をカウント
                m = m + 1
                ' 図のページが最大ページより大きい場合
                If MaxNo < myPictureNo Then
                    ' 最大ページを更新
                    MaxNo = myPictureNo
                End If
            End If
        Next
        ' 写真の最小ページを求める
        MinNo = MaxNo
        ' すべての図にたいして
        For Each myPicture In ActiveSheet.Shapes
            ' 図のページを求める
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' 現在のセルより図のページが大きい場合
            If CurrentNo <= myPictureNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' 図のページが最小ページより小さい場合
                If MinNo > myPictureNo Then
                    ' 最小ページを更新
                    MinNo = myPictureNo
                End If
            End If
        Next
        ' 現在セル以降に写真がない場合
        If MaxNo = 0 Then
            ' 最大ページを現在のページ番号に設定
            MaxNo = CurrentNo
            ' 結合されているセルのカウンタ
            j = 0
            ' くりかえし処理
            Do
                ' 最大ページのセルを選択
                Set myRange = Range(PictureColumn(MaxNo) & Format(PictureRow(MaxNo)))
                myRange.MergeArea.Activate
                ' セルが結合されている場合
                If (ActiveCell.MergeArea.Rows.Count = 10) And (ActiveCell.MergeArea.Columns.Count = 1) Then
                    ' 結合されているセルのカウンタを加算
                    j = j + 1
                    ' 最大ページを加算
                    MaxNo = MaxNo + 1
                    ' 最大ページを超える場合
                    If (j < myFnameCount) And (MaxNo > MaxPageNo) Then
                        MsgBox "写真が最大ページを超えます。" & _
                            vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                ' セルが結合されていない場合
                Else
                    ' 追加するページ数を計算
                    PageInsertCount = ((myFnameCount - j - 1) \ 6) + 1
                    ' 追加するページが最大値を超える場合
                    If (pageNo(ActiveCell.MergeArea.Row, ActiveCell.MergeArea.Column) - 1 + PageInsertCount * 6) _
                        > MaxPageNo Then
                        MsgBox "写真が最大ページを超えます。" & _
                            vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                    ' ページを追加する
                    If AddPages(PageInsertCount) < 0 Then
                        MsgBox "ページを追加できません。" & _
                            vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                End If
            ' セルの結合が取り込み枚数分ない間くりかえし
            Loop While j < myFnameCount
        ' 現在セル以降に写真がある場合
        Else
            ' 余白セルの数
            BlankCount = MinNo - CurrentNo
            ' 余白セルの数が取り込み枚数より少ない場合
            If BlankCount < myFnameCount Then
                ' 結合されているセルのカウンタ
                j = 0
                ' 最大ページを次のコマへ
                MaxNo = MaxNo + 1
                ' くりかえし処理
                Do
                    ' 最大ページを超える場合
                    If (j < (myFnameCount - BlankCount)) And (MaxNo > MaxPageNo) Then
                        MsgBox "写真が最大ページを超えます。" & _
                            vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                   ' 最大ページのセルを選択
                    Set myRange = Range(PictureColumn(MaxNo) & Format(PictureRow(MaxNo)))
                    myRange.MergeArea.Activate
                    ' セルが結合されている場合
                    If (ActiveCell.MergeArea.Rows.Count = 10) And (ActiveCell.MergeArea.Columns.Count = 1) Then
                        ' 結合されているセルのカウンタを加算
                        j = j + 1
                        ' 最大ページを加算
                        MaxNo = MaxNo + 1
                    ' セルが結合されていない場合
                    Else
                        ' 追加するページ数を計算
                        PageInsertCount = ((myFnameCount - j - BlankCount - 1) \ 6) + 1
                        ' 追加するページが最大値を超える場合
                        If (pageNo(ActiveCell.MergeArea.Row, ActiveCell.MergeArea.Column) - 1 + _
                            PageInsertCount * 6) > MaxPageNo Then
                            MsgBox "写真が最大ページを超えます。" & _
                                vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
                            Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                            Exit Sub
                        End If
                        ' ページを追加する
                        If AddPages(PageInsertCount) < 0 Then
                            MsgBox "ページを追加できません。" & _
                                vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
                            Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                            Exit Sub
                        End If
                    End If
                ' セルが結合されている数が取り込み枚数分ない間くりかえし
                Loop While j < (myFnameCount - BlankCount)
                ' ずらす写真枚数が１枚以上ある場合
                If m > 0 Then
                    Application.StatusBar = "写真をずらしています。お待ちください。"
                    ' 配列変数を宣言
                    ReDim PictureName(m - 1)
                    ' 配列のカウンタ
                    j = 0
                    ' すべての図に対して
                    For Each myPicture In ActiveSheet.Shapes
                        ' 写真の通し番号を求める
                        myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
                        ' 現在の通し番号より大きい場合
                        If myPictureNo >= CurrentNo And _
                            (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                            ' 写真の名前を配列変数に格納
                            PictureName(j) = myPicture.Name
                            ' 配列のカウンタを加算
                            j = j + 1
                        End If
                    Next
                    ' 写真の枚数が２枚以上なら
                    If m > 1 Then
                        ' 配列の昇順ならべかえ
                        Call QuickSort(PictureName, 0, m - 1)
                    End If
                    ' 写真をずらす
                    For j = m - 1 To 0 Step -1
                        ' 写真のページ番号を計算
                        myPictureNo = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
                                            ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
                        ' ずらすページ番号を計算
                        NextNo = myPictureNo + myFnameCount - BlankCount
                        ' ずらすページの行と列を求める
                        NextRow = PictureRow(NextNo)
                        NextColumn = PictureColumn(NextNo)
                        With ActiveSheet.Shapes(PictureName(j))
                            ' 写真をずらす
                            .Top = Range(NextColumn & Format(NextRow)).Top
                            .Left = Range(NextColumn & Format(NextRow)).Left
                            Range(NextColumn & Format(NextRow)).MergeArea.Select
                            ' 位置決め
                            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
                            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
                        End With
                        ' 撮影日データをずらす
                        Range(PictureNameBuffer & Format(MinDataBuffer + NextNo)).Value = _
                            Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value
                        Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value = ""
                        Range(PictureDateBuffer & Format(MinDataBuffer + NextNo)).Value = _
                            Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value
                        Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value = ""
                        ' 撮影日表示フラグがＯＮなら
                        If Range(PictureDateFlag).Value <> 0 Then
                            ' 撮影日テキストボックスをずらす
                            With ActiveSheet.Shapes(PictureDateName(PictureName(j)))
                                .Top = Range(NextColumn & Format(NextRow)).Top
                                .Left = Range(NextColumn & Format(NextRow)).Left
                                Range(NextColumn & Format(NextRow)).MergeArea.Select
                            End With
                            Call PictureDatePosition(PictureName(j))
                        End If
                        ' コメントのずらす行と列を計算
                        CurrentRow = PictureRow(myPictureNo)
                        CurrentColumn = CommentColumn(myPictureNo)
                        NextColumn = CommentColumn(NextNo)
                        ' コメントをずらす
                        Range(NextColumn & Format(NextRow + 1) & ":" & NextColumn & Format(NextRow + 9)).Value = _
                            Range(CurrentColumn & Format(CurrentRow + 1) & _
                                ":" & CurrentColumn & Format(CurrentRow + 9)).Value
                        ' ずらしたコメントを消去
                        Range(CurrentColumn & Format(CurrentRow + 1) & _
                            ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
                    Next j
                    ' 配列変数の解放
                    Erase PictureName
                End If
            End If
        End If
        ' Excel2007以降で、画像がずれる対策
        myCurrentWindowZoom = ActiveWindow.Zoom
        ' ウィンドウの表示倍率を１００％にする
        ActiveWindow.Zoom = 100
        Application.StatusBar = "写真を取り込んでいます。お待ちください。"
        ' 写真の取り込み枚数分くりかえし
       For i = 1 To UBound(SelectedOrder)
    If SelectedOrder(i) > 0 Then
        Set targetCell = PasteStartCell.Offset(SelectedOrder(i) - 1, 0)
        Set myRange = Range(PictureColumn(CurrentNo) & PictureRow(CurrentNo))
        ' Activate不要なら削除
        Call PastePicture(SelectedPaths(i))
        CurrentNo = CurrentNo + 1
    End If
Next i
        ' 配列変数の解放
        Erase myFname
        ' Excel2007以降で、画像がずれる対策
        ' ウィンドウの表示倍率を復元する
        ActiveWindow.Zoom = myCurrentWindowZoom
    End If
    ' 配列変数の解放
    'Erase myFilename
    ' 最初のセルを選択
    Set myRange = Range(PictureColumn(StartNo) & Format(PictureRow(StartNo)))
    myRange.MergeArea.Select
    ' 同じ名前の写真カウンタが「０」以外なら
    If myFailureCount > 0 Then
        ' 確認メッセージを表示
        If myFailureCount = mySelectedItemsCount Then
            MsgBox "同じ写真を取り込むことはできません。", vbOKOnly + vbExclamation, "お知らせ"
        Else
            MsgBox Format(mySelectedItemsCount) & "枚中" & Format(myFailureCount) & "枚同じ写真がありました。" _
            & vbCrLf & "同じ写真は取り込まれていません。確認してください。", vbOKOnly + vbExclamation, "お知らせ"
        End If
    End If
    ' 長いファイル名の写真カウンタが「０」以外なら
    If myLongNameCount > 0 Then
        ' 確認メッセージを表示
        If myLongNameCount = mySelectedItemsCount Then
            MsgBox "写真のファイル名が長すぎます" _
                & vbCrLf & "ファイル名は拡張子（.JPGなど）を含めて３０字以下としてください。", _
                vbOKOnly + vbExclamation, "お知らせ"
        Else
            MsgBox Format(mySelectedItemsCount) & "枚中" & Format(myLongNameCount) & _
                "枚長いファイル名の写真がありました。" _
                & vbCrLf & "ファイル名は拡張子（.JPGなど）を含めて３０字以下としてください。" _
                & vbCrLf & "長いファイル名の写真は取り込まれていません。確認してください。", _
            vbOKOnly + vbExclamation, "お知らせ"
        End If
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function AddBlankFunc() As Integer
' コマの追加処理
    Dim i As Integer, j As Integer
    Dim CurrentNo As Integer, PictureNo As Integer, MaxNo As Integer, NextNo As Integer
    Dim CurrentRow As Long, CurrentColumn As String
    Dim NextRow As Long, NextColumn As String
    Dim MaxRow As Long, MaxColumn As String
    Dim myPicture As Shape
    Dim PictureName() As String
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        ' 処理を終了
        AddBlankFunc = -1
        Exit Function
    End If
    ' 行、列から、現在のページ番号を求める
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' 現在のページ番号が「０」以下の場合
    If CurrentNo <= 0 Then
        ' 処理を終了
        AddBlankFunc = -3
        Exit Function
    End If
    Application.ScreenUpdating = False
    ' 写真の最大ページ番号を求める
    i = 0
    MaxNo = 0
    ' すべての図に対して
    For Each myPicture In ActiveSheet.Shapes
        ' 最小セル以降の図に対して
        If myPicture.TopLeftCell.Row >= myMinRow Then
            PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' 写真の通し番号が現在の写真より大きい場合
            If PictureNo >= CurrentNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' 写真の枚数をカウント
                i = i + 1
                ' 写真の最大ページ番号を更新
                If MaxNo < PictureNo Then
                    MaxNo = PictureNo
                End If
            End If
        End If
    Next
    ' 写真の最大ページ番号の次にする
    MaxNo = MaxNo + 1
    ' 写真が最大ページを超える場合
    If MaxNo > MaxPageNo Then
        ' 処理を終了
        AddBlankFunc = -4
        Application.ScreenUpdating = True
        Exit Function
    End If
    ' 写真の最大ページの行と列を求める
    MaxRow = PictureRow(MaxNo)
    MaxColumn = PictureColumn(MaxNo)
    ' ページを追加する
    If (Range(MaxColumn & Format(MaxRow)).MergeArea.Rows.Count <> 10) Or _
        (Range(MaxColumn & Format(MaxRow)).MergeArea.Columns.Count <> 1) Then
        ' ページを追加できなければ
        If AddPages(1) < 0 Then
            ' 処理を終了
            AddBlankFunc = -2
            Application.ScreenUpdating = True
            Exit Function
        End If
    End If
    ' 写真がある場合
    If i > 0 Then
        ' ステータスバーに状態を表示する
        Application.StatusBar = "写真をずらしています。お待ちください。"
        ' 配列変数を宣言
        ReDim PictureName(i - 1)
        j = 0
        ' すべての図に対して
        For Each myPicture In ActiveSheet.Shapes
            ' 最小の行以降の図形に対して
            If myPicture.TopLeftCell.Row >= myMinRow Then
                ' 写真の通し番号を求める
                PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
                ' 現在の通し番号より大きい場合
                If PictureNo >= CurrentNo And _
                    (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                    ' 写真の名前を配列変数に格納
                    PictureName(j) = myPicture.Name
                    ' 配列のカウンタを加算
                    j = j + 1
                End If
            End If
        Next
        ' 写真の枚数が２枚以上なら
        If i > 1 Then
            ' 配列の昇順ならべかえ
            Call QuickSort(PictureName, 0, i - 1)
        End If
        ' 写真をずらす
        For j = i - 1 To 0 Step -1
            ' 写真の行と列から写真の通し番号を求める
            PictureNo = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
                                ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
            ' ずらす写真の番号と行および列を求める
            NextNo = PictureNo + 1
            NextRow = PictureRow(NextNo)
            NextColumn = PictureColumn(NextNo)
            With ActiveSheet.Shapes(PictureName(j))
                ' 写真をずらす
                .Top = Range(NextColumn & Format(NextRow)).Top
                .Left = Range(NextColumn & Format(NextRow)).Left
                ' ずらしたセルを選択
                Range(NextColumn & Format(NextRow)).MergeArea.Select
                ' 写真の位置決め
                .Top = Selection.Top + ((Selection.Height - .Height) / 2)
                .Left = Selection.Left + ((Selection.Width - .Width) / 2)
            End With
            ' 撮影日表示フラグがＯＮなら
            If Range(PictureDateFlag).Value <> 0 Then
                ' 撮影日テキストボックスをずらす
                With ActiveSheet.Shapes(PictureDateName(PictureName(j)))
                    .Top = Range(NextColumn & Format(NextRow)).Top
                    .Left = Range(NextColumn & Format(NextRow)).Left
                    Range(NextColumn & Format(NextRow)).MergeArea.Select
                End With
                Call PictureDatePosition(PictureName(j))
            End If
            ' 撮影日データをずらす
            Range(PictureNameBuffer & Format(MinDataBuffer + NextNo)).Value = _
                Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value
            Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
            Range(PictureDateBuffer & Format(MinDataBuffer + NextNo)).Value = _
                Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value
            Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
            ' コメントの現在の行と列を求める
            CurrentRow = PictureRow(PictureNo)
            CurrentColumn = CommentColumn(PictureNo)
            ' ずらすコメントの列を求める
            NextColumn = CommentColumn(NextNo)
            ' コメントをずらす
            Range(NextColumn & Format(NextRow + 1) & ":" & NextColumn & Format(NextRow + 9)).Value = _
                Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value
            ' ずらしたセルのコメントを消去
            Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
        Next j
        ' 配列変数の解放
        Erase PictureName
    End If
    ' 正常終了
    AddBlankFunc = 0
    ' 追加した余白セルを選択
    Range(PictureColumn(CurrentNo) & Format(PictureRow(CurrentNo))).MergeArea.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Function

Sub AddBlank()
' コマの追加処理
    Select Case AddBlankFunc()
        ' 戻り値による処理
        Case -1
            MsgBox "シートが保護されています。" _
                & vbCrLf & "保護を解除してください。", vbOKOnly + vbExclamation, "お知らせ"
        Case -2
            MsgBox "これ以上ページを追加できません。", vbOKOnly + vbExclamation, "お知らせ"
        Case -3
            MsgBox "コマ番号が正しくありません。", vbOKOnly + vbExclamation, "お知らせ"
        Case -4
            MsgBox "写真が最大ページを超えます。" _
                & vbCrLf & "処理を中断します。", vbOKOnly + vbExclamation, "お知らせ"
    End Select
End Sub

Sub DeleteBlank()
' コマの削除処理
    Dim CurrentNo As Integer, PictureNo As Integer, NextNo As Integer
    Dim CurrentRow As Long, CurrentColumn As String
    Dim NextRow As Long, NextColumn As String
    Dim i As Integer, j As Integer
    Dim myPicture As Shape
    Dim PictureName() As String
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        MsgBox "シートが保護されています。" _
            & vbCrLf & "保護を解除してください。", vbOKOnly + vbExclamation, "お知らせ"
        ' 処理を終了
        Exit Sub
    End If
    ' 行、列から、現在のページ番号を求める
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' 現在のページ番号が「０」以下の場合
    If CurrentNo <= 0 Then
        ' 処理を終了
        Exit Sub
    End If
    Application.ScreenUpdating = False
    ' 写真の枚数を求める
    i = 0
    For Each myPicture In ActiveSheet.Shapes
        PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        If PictureNo > CurrentNo And _
            (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
            i = i + 1
        End If
    Next
    ' 削除するコマの写真枚数がゼロなら
    If i = 0 Then
        ' 現在のセルのコメントを消去
        CurrentRow = PictureRow(CurrentNo)
        CurrentColumn = CommentColumn(CurrentNo)
        Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
        ' 撮影日データを消去
        Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
        Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Application.StatusBar = "写真をずらしています。お待ちください。"
    ' 配列変数を宣言
    ReDim PictureName(i - 1)
    j = 0
    ' すべての図に対して
    For Each myPicture In ActiveSheet.Shapes
        ' 図のページ番号を求める
        PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' 写真が現在のページ番号より大きい場合
        If PictureNo > CurrentNo And _
            (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
            ' 写真の名前を配列変数に格納
            PictureName(j) = myPicture.Name
            ' 配列のカウンタを加算
            j = j + 1
        End If
    Next
    ' 写真の枚数が２枚以上なら
    If i > 1 Then
        ' 配列の昇順ならべかえ
        Call QuickSort(PictureName, 0, i - 1)
    End If
    ' 写真をずらす
    For j = 0 To i - 1
        PictureNo = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
                            ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
        ' ずらす写真の通し番号と行と列を求める
        NextNo = PictureNo - 1
        NextRow = PictureRow(NextNo)
        NextColumn = PictureColumn(NextNo)
        With ActiveSheet.Shapes(PictureName(j))
            ' 写真をずらす
            .Top = Range(NextColumn & Format(NextRow)).Top
            .Left = Range(NextColumn & Format(NextRow)).Left
            Range(NextColumn & Format(NextRow)).MergeArea.Select
            ' 「Ａ２」セルで写真が縦長で９０°または２７０°の場合に位置決めがずれる対策
            If NextRow = 2 Then
                Range("A1").RowHeight = TempRowHeight
            End If
            ' 写真の位置決め
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
            ' 「Ａ２」セルの写真が縦長で９０°または２７０°の場合に位置決めがずれる対策
            If NextRow = 2 Then
                Range("A1").RowHeight = TopRowHeight
            End If
        End With
        ' 撮影日表示フラグがＯＮなら
        If Range(PictureDateFlag).Value <> 0 Then
            ' 撮影日テキストボックスをずらす
            With ActiveSheet.Shapes(PictureDateName(PictureName(j)))
                .Top = Range(NextColumn & Format(NextRow)).Top
                .Left = Range(NextColumn & Format(NextRow)).Left
                Range(NextColumn & Format(NextRow)).MergeArea.Select
            End With
            Call PictureDatePosition(PictureName(j))
        End If
        ' 撮影日データをずらす
        Range(PictureNameBuffer & Format(MinDataBuffer + NextNo)).Value = _
            Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value
        Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
        Range(PictureDateBuffer & Format(MinDataBuffer + NextNo)).Value = _
            Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value
        Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
        ' コメントの現在の行と列を求める
        CurrentRow = PictureRow(PictureNo)
        CurrentColumn = CommentColumn(PictureNo)
        ' ずらすコメントの列を求める
        NextColumn = CommentColumn(NextNo)
        ' コメントをずらす
        Range(NextColumn & Format(NextRow + 1) & ":" & NextColumn & Format(NextRow + 9)).Value = _
            Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value
        ' ずらしたセルのコメントを消去
        Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
    Next j
    ' 配列変数の解放
    Erase PictureName
    ' 削除したセルを選択
    Range(PictureColumn(CurrentNo) & Format(PictureRow(CurrentNo))).MergeArea.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function pageNo(myRow As Long, myColumn As Long) As Integer
' 行番号、列番号から、写真の通し番号を求める
    If myColumn = 1 Then
        pageNo = (myRow - myMinRow) \ 11 + 1
    Else
        pageNo = 0
    End If
End Function

Function PictureRow(PictureNo As Integer) As Long
' 写真の通し番号から、行番号を求める
    PictureRow = CLng(PictureNo - 1) * 11 + myMinRow
End Function

Function PictureColumn(PictureNo As Integer) As String
' 写真の通し番号から、列番号を求める
    PictureColumn = "A"
End Function

Function CommentColumn(PictureNo As Integer) As String
' 写真の通し番号から、コメント欄の列番号を求める
    CommentColumn = "B"
End Function

Sub QuickSort(ByRef PictureName() As String, ByVal ArrayMinNo As Integer, ByVal ArrayMaxNo As Integer)
' 配列の並べ替え処理（クイックソート）
    Dim MinNo As Integer
    Dim MaxNo As Integer
    Dim MidValue As Integer
    Dim TempValue As String
    ' 中央の値を取得
    MidValue = pageNo(ActiveSheet.Shapes(PictureName(Int((ArrayMinNo + ArrayMaxNo) / 2))).TopLeftCell.Row, _
                ActiveSheet.Shapes(PictureName(Int((ArrayMinNo + ArrayMaxNo) / 2))).TopLeftCell.Column)
    MinNo = ArrayMinNo
    MaxNo = ArrayMaxNo
    ' くりかえし
    Do
        ' 配列の最小側のくりかえし
        Do
            ' 配列の値が中央の値より大きい場合
            If (pageNo(ActiveSheet.Shapes(PictureName(MinNo)).TopLeftCell.Row, _
                    ActiveSheet.Shapes(PictureName(MinNo)).TopLeftCell.Column) >= MidValue) Then
                ' くりかえしを抜ける
                Exit Do
            End If
            ' 最小側を１つ加算
            MinNo = MinNo + 1
        Loop
        ' 配列の最大側のくりかえし
        Do
            ' 配列の値が中央の値より小さい場合
            If (MidValue >= pageNo(ActiveSheet.Shapes(PictureName(MaxNo)).TopLeftCell.Row, _
                            ActiveSheet.Shapes(PictureName(MaxNo)).TopLeftCell.Column)) Then
                ' くりかえしを抜ける
                Exit Do
            End If
            ' 最大側を１つ減算
            MaxNo = MaxNo - 1
        Loop
        ' 最小側が最大側より大きくなれば処理終了
        If (MinNo >= MaxNo) Then
            Exit Do
        End If
        ' 配列の入れ替え
        TempValue = PictureName(MinNo)
        PictureName(MinNo) = PictureName(MaxNo)
        PictureName(MaxNo) = TempValue
        ' 最小側を１つ加算
        MinNo = MinNo + 1
        ' 最大側を１つ減算
        MaxNo = MaxNo - 1
    Loop
    ' 配列の最小側を再帰で並べ替え
    If (ArrayMinNo < MinNo - 1) Then
        Call QuickSort(PictureName, ArrayMinNo, MinNo - 1)
    End If
    ' 配列の最大側を再帰で並べ替え
    If (MaxNo + 1 < ArrayMaxNo) Then
        Call QuickSort(PictureName, MaxNo + 1, ArrayMaxNo)
    End If
End Sub

Sub PictureNumbering()
' 写真毎の番号付番処理
    Dim myPage As Integer, i As Integer, j As Integer, k As Integer
    Dim myPicture As Shape, PictureName() As String, PictureNo As Integer
    Dim myRange As Range, CurrentRange As Range, CurrentRow As Long, CurrentColumn As Long
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        MsgBox "シートが保護されています。" _
            & vbCrLf & "保護を解除してください。", vbOKOnly + vbExclamation, "お知らせ"
        ' 処理を終了
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.StatusBar = "写真毎に番号を振り直しています。お待ちください。"
    ' 現在のセルを記憶
    Set CurrentRange = Cells(ActiveCell.Row, ActiveCell.Column)
    ' 写真の枚数カウンタ
    i = 0
    ' 写真の枚数を求める
    For Each myPicture In ActiveSheet.Shapes
        PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' 図が写真でページ番号が「０」より大きい場合
        If PictureNo > 0 And (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
            ' 枚数をカウント
            i = i + 1
        End If
    Next
    ' 写真が１枚以上ある場合
    If i > 0 Then
        ' 配列変数を宣言
        ReDim PictureName(i - 1)
        ' 配列変数用カウンタ
        j = 0
        ' すべての図に対して
        For Each myPicture In ActiveSheet.Shapes
            PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' 図が写真でページ番号が「０」より大きい場合
            If PictureNo > 0 And (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' 配列に図の名前を代入
                PictureName(j) = myPicture.Name
                ' カウンタを加算
                j = j + 1
            End If
        Next
        ' 写真が２枚以上なら
        If i > 1 Then
            ' 配列の並べ替え
            Call QuickSort(PictureName, 0, i - 1)
        End If
    End If
    ' 写真の通し番号
    j = 0
    ' コマの通し番号
    k = 1
    ' ページ番号
    myPage = 1
    ' 先頭のコマのセルを選択
    Set myRange = Range(PictureColumn(1) & Format(PictureRow(1)))
    myRange.Select
    ' くりかえし
    Do
        ' 写真が１枚以上の場合
        If i > 0 And j < i Then
            ' 写真の行と列を求める
            CurrentRow = ActiveSheet.Shapes(PictureName(j)).TopLeftCell.MergeArea.Row
            CurrentColumn = ActiveSheet.Shapes(PictureName(j)).TopLeftCell.MergeArea.Column
            ' 現在のセルの行と列と写真の行と列が一致した場合
            If myRange.Row = CurrentRow And myRange.Column = CurrentColumn Then
                ' ページ番号をセルに書き込む
                Range(CommentColumn(k) & Format(PictureRow(k))).Value = "No." & Format(myPage)
                 ' ページを加算
                myPage = myPage + 1
               ' つぎの写真へ
                j = j + 1
            ' 一致しない場合
            Else
                ' ページ番号を消去
                Range(CommentColumn(k) & Format(PictureRow(k))).Value = ""
            End If
        ' 最終ページまで
        Else
            ' ページ番号を消去
            Range(CommentColumn(k) & Format(PictureRow(k))).Value = ""
        End If
        ' 次のコマへ移動
        k = k + 1
        Set myRange = Range(PictureColumn(k) & Format(PictureRow(k)))
        myRange.MergeArea.Select
    ' セルが結合している間くりかえし
    Loop While (myRange.MergeArea.Rows.Count = 10) And (k <= MaxPageNo)
    ' 配列変数の解放
    Erase PictureName
    ' 現在のセルを選択
    CurrentRange.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub SerialNumbering()
' コマごとの番号付番処理
    Dim myPage As Integer, k As Integer
    Dim myRange As Range, CurrentRange As Range
    ' シートが保護されている場合
    If ActiveSheet.ProtectContents Then
        MsgBox "シートが保護されています。" _
            & vbCrLf & "保護を解除してください。", vbOKOnly + vbExclamation, "お知らせ"
        ' 処理を終了
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.StatusBar = "コマ毎に番号を振り直しています。お待ちください。"
    ' 現在のセルを記憶
    Set CurrentRange = Cells(ActiveCell.Row, ActiveCell.Column)
    ' コマの通し番号
    k = 1
    ' ページ番号
    myPage = 1
    ' 先頭のコマのセルを選択
    Set myRange = Range(PictureColumn(1) & Format(PictureRow(1)))
    myRange.Select
    ' くりかえし
    Do
        ' ページ番号をセルに書き込む
        Range(CommentColumn(k) & Format(PictureRow(k))).Value = "No." & Format(myPage)
        ' ページを加算
        myPage = myPage + 1
        ' 次のコマへ移動
        k = k + 1
        Set myRange = Range(PictureColumn(k) & Format(PictureRow(k)))
        myRange.MergeArea.Select
    ' セルが結合している間くりかえし
    Loop While (myRange.MergeArea.Rows.Count = 10) And (k <= MaxPageNo)
    ' 現在のセルを選択
    CurrentRange.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function PictureExist(CurrentRow As Long, CurrentColumn As Long) As Boolean
' セルに写真が貼り付けられているか確認する関数
    Dim CurrentNo As Integer, PictureNo As Integer
    Dim myPicture As Shape
    ' 現在のセルのページ番号を求める
    CurrentNo = pageNo(CurrentRow, CurrentColumn)
    ' すべての写真について繰り返し
    For Each myPicture In ActiveSheet.Shapes
        If myPicture.TopLeftCell.Row >= myMinRow Then
            ' 写真のページ番号を求める
            PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' 写真のページ番号と現在のセルのページ番号が等しければ
            If PictureNo = CurrentNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' 戻り値を「True」にする
                PictureExist = True
                ' 処理を終了
                Exit Function
            End If
        End If
    Next
    ' 現在のセルに写真が貼り付けられていなければ戻り値を「False」にする
    PictureExist = False
End Function

Sub MoveToEnd()
' 写真を末尾に移動させる処理
    Dim myPicture As Shape
    Dim myRange As Range
    Dim myPictureNo As Integer, myMaxNo As Integer
    Dim CommentData(10) As String, i As Integer, AdjustNo As Integer
    Dim myPictureName As String, myPictureDate As String, myDate As String, myType As Integer
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' 図形の最大ページ数
    myMaxNo = 0
    ' 同じ名前のカウンタ
    i = 0
    ' 移動先と移動元のシートが異なれば
    If SwapSourceSheet <> SwapDestSheet Then
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
    End If
    ' すべての図形に対して
    For Each myPicture In ActiveSheet.Shapes
        ' 図形の左上セルが最小セル以上の場合
        If myPicture.TopLeftCell.Row >= myMinRow Then
            ' 図形のページ数を求める
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' ページの最大数を求める
            If myMaxNo < myPictureNo Then
                myMaxNo = myPictureNo
            End If
        End If
        ' 移動先と移動元のシートが異なれば
        If SwapSourceSheet <> SwapDestSheet Then
            ' 写真の名前が移動元の写真の名前と同じなら
            If myPicture.Name = SwapSourceName Then
                ' カウンタを加算
                i = i + 1
                ' カウンタが１以上なら
                If i > 0 Then
                    ' メッセージを表示して処理を終了
                    MsgBox "移動先のシートに同じ写真があります。" & vbCrLf & "同じ写真を貼り付けることはできません。", _
                        vbOKOnly + vbExclamation, "お知らせ"
                    Exit Sub
                End If
            End If
        End If
    Next
    ' 移動先のシートを選択
    Worksheets(SwapDestSheet).Select
    ' 移動元と移動先のシートが異なる場合
    If SwapSourceSheet <> SwapDestSheet Then
        ' ページ番号の補正値を＋１
        AdjustNo = 1
        ' 最大ページのセルが結合されていなければ
        If Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo))).MergeArea.Rows.Count <> 10 Or _
            Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo))).MergeArea.Columns.Count <> 1 Then
            ' ページを追加する
            If AddPages(1) < 0 Then
                ' ページの追加ができない場合メッセージを表示して処理を終了
                MsgBox "ページを追加できませんでした。" & vbCrLf & "処理を終了します。", vbOKOnly + vbExclamation, "お知らせ"
                Exit Sub
            End If
        End If
    ' 移動元と移動先のシートが同じ場合
    Else
        ' ページ番号の補正値を０
        AdjustNo = 0
    End If
    ' 移動元シートを選択
    Worksheets(SwapSourceSheet).Select
    ' コメント欄のデータを保存
    For i = 1 To 9
        CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
    Next i
    ' 撮影日データを保存
    myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag).Value <> 0 Then
        ' 撮影日テキストボックスを削除
        ActiveSheet.Shapes(PictureDateName(SwapSourceName)).Delete
    End If
    ' 写真を切り取り
    ActiveSheet.Shapes(SwapSourceName).Cut
    ' コマの削除
    Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
    Call DeleteBlank
    ' 移動先のシートを選択
    Worksheets(SwapDestSheet).Select
    ' 移動元と移動先のシートが異なる場合
    If SwapSourceSheet <> SwapDestSheet Then
        ' ページ番号の補正値を＋１
        AdjustNo = 1
    ' 移動元と移動先のシートが同じ場合
    Else
        ' ページ番号の補正値を０
        AdjustNo = 0
    End If
    ' 最終ページに写真を貼り付け
    Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo))).MergeArea.Select
    ActiveSheet.Paste
    With ActiveSheet.Shapes(SwapSourceName)
        ' 写真を貼り付けたセルを選択
        .TopLeftCell.MergeArea.Select
        ' 写真の縦横比を固定
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
    End With
    ' コメント欄のデータを最大ページに移す
    For i = 1 To 9
        Range(CommentColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo) + i)).Value = CommentData(i)
    Next i
    ' 撮影日データを最大ページに移す
    Range(PictureNameBuffer & Format(MinDataBuffer + myMaxNo + AdjustNo)).Value = myPictureName
    Range(PictureDateBuffer & Format(MinDataBuffer + myMaxNo + AdjustNo)).Value = myPictureDate
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag).Value <> 0 Then
        myDate = myPictureDate
        ' 撮影日の表示形式フラグをエンコード
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        ' 撮影日のテキストボックスを作図
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If
    ' 最大ページのセルを選択
    Set myRange = Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo)))
    myRange.MergeArea.Select
End Sub

Sub MoveToHere()
' 写真を移動させる処理
    Dim SwapSourceRow As Long, SwapSourceColumn As String
    Dim SwapDestRow As Long, SwapDestColumn As String
    Dim CommentData(10) As String, i As Integer
    Dim AdjustNo As Integer
    Dim myPicture As Shape
    Dim myPictureName As String, myPictureDate As String, myDate As String, myType As Integer
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' 写真の移動元の行と列を取得
    SwapSourceRow = PictureRow(SwapSourceNo)
    SwapSourceColumn = PictureColumn(SwapSourceNo)
    ' 写真の移動先の行と列を取得
    SwapDestRow = PictureRow(SwapDestNo)
    SwapDestColumn = PictureColumn(SwapDestNo)
    ' 移動先のシートと移動元のシートが異なる場合
    If SwapSourceSheet <> SwapDestSheet Then
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
        ' 同じ写真のカウンタ
        i = 0
        ' すべての図に対して繰り返し
        For Each myPicture In ActiveSheet.Shapes
            ' 写真の名前が移動元の写真の名前と同じなら
            If myPicture.Name = SwapSourceName Then
                ' カウンタを加算
                i = i + 1
                ' カウンタが１以上なら
                If i > 0 Then
                    ' メッセージを表示して処理を終了
                    MsgBox "移動先のシートに同じ写真があります。" & vbCrLf & "同じ写真を貼り付けることはできません。", _
                        vbOKOnly + vbExclamation, "お知らせ"
                    Exit Sub
                End If
            End If
        Next
        ' 移動先のページを選択
        Range(SwapDestColumn & Format(SwapDestRow)).MergeArea.Select
        ' 余白コマを追加
        If AddBlankFunc < 0 Then
            ' メッセージを表示して処理を終了
            MsgBox "余白コマを追加できませんでした。" & vbCrLf & "処理を終了します。", vbOKOnly + vbExclamation, "お知らせ"
            Exit Sub
        End If
    End If
    ' 移動元のシートを選択
    Worksheets(SwapSourceSheet).Select
    ' 移動元のセルを選択
    ActiveSheet.Shapes(SwapSourceName).TopLeftCell.MergeArea.Select
    ' 移動元のコメント欄のデータを保存
    For i = 1 To 9
        CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
    Next i
    ' 移動元の撮影日データを保存
    myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = ""
    Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = ""
    ' 移動元のセルを選択
    ActiveSheet.Shapes(SwapSourceName).TopLeftCell.MergeArea.Select
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag).Value <> 0 Then
        ' 撮影日テキストボックスを削除
        ActiveSheet.Shapes(PictureDateName(SwapSourceName)).Delete
    End If
    ' 移動元のセルの写真を切り取り
    ActiveSheet.Shapes(SwapSourceName).Cut
    ' 移動元の余白コマを削除
    Call DeleteBlank
    ' 移動先のページ番号が移動元ページ番号より大きい場合
    If (SwapSourceNo < SwapDestNo) And (SwapSourceSheet = SwapDestSheet) Then
        ' 移動先ページ番号の補正値を－１にする
        AdjustNo = -1
    ' 移動先のページ番号が移動元ページ番号より小さい場合
    Else
        ' 移動先ページ番号の補正値を０にする
        AdjustNo = 0
    End If
    ' 移動先と移動元のシートが同じなら
    If SwapSourceSheet = SwapDestSheet Then
        ' 移動先のセルを選択
        Range(PictureColumn(SwapDestNo + AdjustNo) & Format(PictureRow(SwapDestNo + AdjustNo))).MergeArea.Select
        ' 余白コマを追加
        If AddBlankFunc < 0 Then
            ' メッセージを表示して処理を終了
            MsgBox "余白コマを追加できませんでした。" & vbCrLf & "処理を終了します。", vbOKOnly + vbExclamation, "お知らせ"
            Exit Sub
        End If
    ' 移動先と移動元のシートが異なる場合
    Else
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
    End If
    ' 写真の移動先に移動元の写真を貼り付け
    Range(PictureColumn(SwapDestNo + AdjustNo) & Format(PictureRow(SwapDestNo + AdjustNo))).MergeArea.Select
    ActiveSheet.Paste
    With ActiveSheet.Shapes(SwapSourceName)
        ' 写真を貼り付けたセルを選択
        .TopLeftCell.MergeArea.Select
        ' 写真の縦横比を固定
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
    End With
    ' コメント欄のデータを移動先へ移す
    For i = 1 To 9
        Range(CommentColumn(SwapDestNo + AdjustNo) & Format(PictureRow(SwapDestNo + AdjustNo) + i)).Value = CommentData(i)
    Next i
    ' 撮影日データを移動先へ移す
    Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo + AdjustNo)).Value = myPictureName
    Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo + AdjustNo)).Value = myPictureDate
    ' 撮影日表示フラグがＯＮなら
    If Range(PictureDateFlag).Value <> 0 Then
        myDate = myPictureDate
        ' 撮影日の表示形式フラグをエンコード
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        ' 撮影日のテキストボックスを作図
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If
    ' 配列変数の解放
    Erase CommentData
End Sub

Sub ExchangePicture()
' 写真を入れ替える処理
    Dim SwapSourceRow As Long, SwapSourceColumn As String
    Dim SwapDestRow As Long, SwapDestColumn As String
    Dim CommentData(10) As String, i As Integer
    Dim myPicture As Shape
    Dim myPictureName As String, myPictureDate As String, myDestPictureName As String, myDestPictureDate As String
    Dim SwapSourceDate As String, SwapDestDate As String, myDate As String, myType As Integer
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' 写真の入れ替え元の行と列を取得
    SwapSourceRow = PictureRow(SwapSourceNo)
    SwapSourceColumn = PictureColumn(SwapSourceNo)
    ' 写真の入れ替え先の行と列を取得
    SwapDestRow = PictureRow(SwapDestNo)
    SwapDestColumn = PictureColumn(SwapDestNo)
    ' 移動先と移動元のシートが同じ場合
    If SwapSourceSheet = SwapDestSheet Then
        ' 写真の移動先に移動元の写真を移動
        With ActiveSheet.Shapes(SwapSourceName)
            .Top = Range(SwapDestColumn & Format(SwapDestRow)).Top
            .Left = Range(SwapDestColumn & Format(SwapDestRow)).Left
            Range(SwapDestColumn & Format(SwapDestRow)).Select
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        End With
        ' 写真の移動元に移動先の写真を移動
        With ActiveSheet.Shapes(SwapDestName)
            .Top = Range(SwapSourceColumn & Format(SwapSourceRow)).Top
            .Left = Range(SwapSourceColumn & Format(SwapSourceRow)).Left
            Range(SwapSourceColumn & Format(SwapSourceRow)).Select
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        End With
        ' 移動元のコメント欄のデータを取得
        For i = 1 To 9
            CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
        Next i
        ' コメント欄のデータを移動先から移動元に移す
        Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 1) & ":" & _
            CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 9)).Value = _
        Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 1) & ":" & _
            CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 9)).Value
        ' コメント欄のデータを移動先へ移す
        For i = 1 To 9
            Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + i)).Value = CommentData(i)
        Next i
        ' 移動元の撮影日データを取得
        myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        ' 撮影日データを移動先から移動元へ移す
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = _
            Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = _
            Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        ' 撮影日データを移動先へ移す
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureName
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureDate
        SwapSourceDate = PictureDateName(SwapSourceName)
        SwapDestDate = PictureDateName(SwapDestName)
        ' 撮影日表示フラグがＯＮなら
        If Range(PictureDateFlag).Value <> 0 Then
            ' 写真の移動先に移動元の撮影日を移動
            With ActiveSheet.Shapes(SwapSourceDate)
                .Top = Range(SwapDestColumn & Format(SwapDestRow)).Top
                .Left = Range(SwapDestColumn & Format(SwapDestRow)).Left
                Range(SwapDestColumn & Format(SwapDestRow)).Select
            End With
            Call PictureDatePosition(SwapSourceName)
            ' 写真の移動元に移動先の撮影日を移動
            With ActiveSheet.Shapes(SwapDestDate)
                .Top = Range(SwapSourceColumn & Format(SwapSourceRow)).Top
                .Left = Range(SwapSourceColumn & Format(SwapSourceRow)).Left
                Range(SwapSourceColumn & Format(SwapSourceRow)).Select
            End With
            Call PictureDatePosition(SwapDestName)
        End If
    ' 移動先と移動元のシートが異なる場合
    Else
        ' 移動元のシートを選択
        Worksheets(SwapSourceSheet).Select
        ' 同じ写真のカウンタ
        i = 0
        ' すべての図に対して繰り返し
        For Each myPicture In ActiveSheet.Shapes
            ' 写真の名前が移動先の写真の名前と同じなら
            If myPicture.Name = SwapDestName Then
                ' カウンタを加算
                i = i + 1
                ' カウンタが１以上なら
                If i > 0 Then
                    ' メッセージを表示して処理を終了
                    MsgBox "移動先のシートに同じ写真があります。" & vbCrLf & "同じ写真を貼り付けることはできません。", _
                        vbOKOnly + vbExclamation, "お知らせ"
                    Exit Sub
                End If
            End If
        Next
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
        ' 同じ写真のカウンタ
        i = 0
        ' すべての図に対して繰り返し
        For Each myPicture In ActiveSheet.Shapes
            ' 写真の名前が移動先の写真の名前と同じなら
            If myPicture.Name = SwapSourceName Then
                ' カウンタを加算
                i = i + 1
                ' カウンタが１以上なら
                If i > 0 Then
                    ' メッセージを表示して処理を終了
                    MsgBox "移動先のシートに同じ写真があります。" & vbCrLf & "同じ写真を貼り付けることはできません。", _
                        vbOKOnly + vbExclamation, "お知らせ"
                    Exit Sub
                End If
            End If
        Next
        ' 移動元のシートを選択
        Worksheets(SwapSourceSheet).Select
        ' 移動元、移動先の撮影日データの名前を設定
        SwapSourceDate = PictureDateName(SwapSourceName)
        SwapDestDate = PictureDateName(SwapDestName)
        ' 撮影日表示フラグがＯＮなら
        If Range(PictureDateFlag).Value <> 0 Then
            ' 移動元の撮影日データを削除
            ActiveSheet.Shapes(SwapSourceDate).Delete
        End If
        ' 移動元の写真を切り取り
        ActiveSheet.Shapes(SwapSourceName).Cut
        ' 移動元のコメント欄のデータを取得
        For i = 1 To 9
            CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
        Next i
        ' 移動元の撮影日データを保存
        myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
        ' 写真の移動先に移動元の写真を貼り付け
        Range(PictureColumn(SwapDestNo) & Format(PictureRow(SwapDestNo))).MergeArea.Select
        ActiveSheet.Paste
        With ActiveSheet.Shapes(SwapSourceName)
            ' 写真を貼り付けたセルを選択
            .TopLeftCell.MergeArea.Select
            ' 写真の縦横比を固定
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
        End With
        ' 撮影日表示フラグがＯＮなら
        If Range(PictureDateFlag).Value <> 0 Then
            myDate = myPictureDate
            ' 撮影日の表示形式フラグをエンコード
            myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
            ' 撮影日のテキストボックスを作図
            Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
        End If
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
        ' 撮影日表示フラグがＯＮなら
        If Range(PictureDateFlag).Value <> 0 Then
            ' 移動先の撮影日データを削除
            ActiveSheet.Shapes(SwapDestDate).Delete
        End If
        ' 移動先の写真を切り取り
        ActiveSheet.Shapes(SwapDestName).Cut
        ' コメント欄のデータを移動先から移動元へ移す
        Worksheets(SwapSourceSheet).Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 1) & ":" & _
            CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 9)).Value = _
            Worksheets(SwapDestSheet).Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 1) & ":" & _
            CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 9)).Value
        ' 移動先の撮影日データを保存
        myDestPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        myDestPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        ' 撮影日データを移動先へ移す
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureName
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureDate
        ' 移動元のシートを選択
        Worksheets(SwapSourceSheet).Select
        ' 写真の移動元に移動先の写真を貼り付け
        Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
        ActiveSheet.Paste
        With ActiveSheet.Shapes(SwapDestName)
            ' 写真を貼り付けたセルを選択
            .TopLeftCell.MergeArea.Select
            ' 写真の縦横比を固定
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
        End With
        ' 撮影日表示フラグがＯＮなら
        If Range(PictureDateFlag).Value <> 0 Then
            myDate = myDestPictureDate
            ' 撮影日の表示形式フラグをエンコード
            myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
            ' 撮影日のテキストボックスを作図
            Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
        End If
        ' 撮影日データを移動元へ移す
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = myDestPictureName
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = myDestPictureDate
        ' 移動先のシートを選択
        Worksheets(SwapDestSheet).Select
        ' コメント欄のデータを移動先に移す
        For i = 1 To 9
            Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + i)).Value = CommentData(i)
        Next i
    End If
    ' 配列変数の解放
    Erase CommentData
    ' 移動先のシートを選択
    Worksheets(SwapDestSheet).Select
End Sub

Sub SaveWOMacro()
' マクロなしで保存終了処理
    Dim myWorksheet As Worksheet
    Dim myShape As Shape
    Dim myFname As String, Ans As Integer, mySheet As String
    ' 現在のシート名を取得
    mySheet = ActiveSheet.Name
    ' 現在のブックが変更されている場合
    If ActiveWorkbook.Saved = False Then
        ' 確認メッセージを表示
        Ans = MsgBox("現在のブックは変更されています。" & vbCrLf & "変更を保存しておきますか？", _
            vbYesNoCancel + vbExclamation, "確認")
        ' キャンセルがクリックされた場合
        If Ans = vbCancel Then
            ' 処理を終了
            Exit Sub
        ' ＹＥＳがクリックされた場合
        ElseIf Ans = vbYes Then
            ' 現在のブックを保存
            ActiveWorkbook.Save
        End If
    End If
    ' ブックのファイルネームを取得して拡張子を取り除く
    myFname = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    ' 繰り返し
    Do
        ' ファイル名の入力
        myFname = InputBox("ファイル名を入力してください。", "マクロを削除してファイルを保存し終了します。", myFname)
        ' キャンセルがクリックされた場合
        If myFname = "" Then
            ' 処理を終了
            Exit Sub
        End If
        ' ファイル名のチェック
        If InStr(myFname, "*") > 0 Or InStr(myFname, "?") > 0 Or InStr(myFname, "\") > 0 Or _
            InStr(myFname, ":") > 0 Or InStr(myFname, "<") > 0 Or InStr(myFname, ">") > 0 Or _
            InStr(myFname, "[") > 0 Or InStr(myFname, "]") > 0 Or InStr(myFname, "|") > 0 Or _
            InStr(myFname, Chr(34)) > 0 Then
            ' メッセージを表示
            Ans = MsgBox("ファイル名が不正です。以下の文字は使えません。" & vbCrLf & _
                " * ? : < > [ ] | \ " & Chr(34), vbOKOnly + vbCritical, "注意")
        ' ファイル名の文字数チェック
        ElseIf Len(ActiveWorkbook.Path & "\" & myFname & ".xlsx") > 218 Then
            ' メッセージを表示
            Ans = MsgBox("ファイル名が長すぎます。", vbOKOnly + vbCritical, "注意")
        ' 同じファイル名のブックがある場合
        ElseIf Dir(ActiveWorkbook.Path & "\" & myFname & ".xlsx") <> "" Then
            ' メッセージを表示
            Ans = MsgBox(myFname & ".xlsxは既に存在します。" & vbCrLf & "上書きしますか？", _
                vbYesNoCancel + vbExclamation, "確認")
        Else
            Ans = vbYes
        End If
        ' キャンセルがクリックされた場合
        If Ans = vbCancel Then
            ' 処理を終了
            Exit Sub
        End If
    ' ＹＥＳをクリックするまで繰り返し
    Loop While Ans <> vbYes
    ' ブック内のすべてのワークシートについて繰り返し
    For Each myWorksheet In ThisWorkbook.Worksheets
        ' ワークシートを選択
        myWorksheet.Select
        ' ワークシート内のすべての図形について繰り返し
        For Each myShape In ActiveSheet.Shapes
            ' 図形のマクロを削除
            If myShape.OnAction <> "" Then
                myShape.OnAction = ""
            End If
        Next
    Next
    ' シートを選択
    Worksheets(mySheet).Select
    ' マクロなしでファイルを保存
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & myFname, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    ' エクセルを終了
    Application.Quit
End Sub

Function PictureDate(myPic As String) As String
' 写真の撮影日時を取得する
    Dim objFS, objFile, shellObj, folderObj, MyFile, myPath, i As Long, GetDetailsNumber As Long

    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFS.GetFile(myPic)
    MyFile = objFile.Name
    myPath = Replace(myPic, MyFile, "")
    myPath = Left(myPath, Len(myPath) - 1)
    Set shellObj = CreateObject("Shell.Application")
    Set folderObj = shellObj.Namespace(myPath)

    GetDetailsNumber = -1
    ' Exif情報の項目番号を取得
    For i = 0 To 100 '十分に大きい数字
        If folderObj.GetDetailsOf("", i) = "撮影日時" Then
            GetDetailsNumber = i
            Exit For
        End If
    Next i
    ' 撮影日時を取得し、戻り値にする
    PictureDate = folderObj.GetDetailsOf(folderObj.ParseName(MyFile), GetDetailsNumber)
    
    Set objFS = Nothing
    Set objFile = Nothing
    Set shellObj = Nothing
    Set folderObj = Nothing
End Function

Sub PictureDateOFF()
' 撮影日時消去処理
    Dim i As Integer, j As Integer, k As Integer, myShape As Shape, PictureName() As String
    ' エラーが発生したら次の処理へ
    On Error Resume Next
    ' 写真の枚数を求める
    i = 0
    For Each myShape In ActiveSheet.Shapes
        If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
            i = i + 1
        End If
    Next
    ' 写真が１枚以上なら
    If i > 0 Then
        ' 動的配列の宣言
        ReDim PictureName(i - 1)
        j = 0
        ' すべての図に対して
        For Each myShape In ActiveSheet.Shapes
            ' 図が写真なら
            If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
                ' 写真の名前を取得
                PictureName(j) = myShape.Name
                ' 写真の枚数をカウント
                j = j + 1
            End If
        Next
        ' 全ての図に対して
        For k = 0 To j - 1
            ' 撮影日テキストボックスを削除
            ActiveSheet.Shapes(PictureDateName(PictureName(k))).Delete
        Next k
    End If
    ' 撮影日表示フラグをリセット
    Range(PictureDateFlag).Value = ""
    Erase PictureName
End Sub

Sub PictureDateON()
' 撮影日時表示処理
    Dim i As Integer, j As Integer, k As Integer, myShape As Shape, PictureName() As String, myDate As String
    Dim CurrentRow As Long, CurrentColumn As Long, myType As Integer
    CurrentRow = ActiveCell.Row
    CurrentColumn = ActiveCell.Column
    ' 写真の枚数を求める
    i = 0
    For Each myShape In ActiveSheet.Shapes
        If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
            i = i + 1
        End If
    Next
    ' 写真が１枚以上なら
    If i > 0 Then
        ' 動的配列の宣言
        ReDim PictureName(i - 1)
        j = 0
        ' すべての図に対して
        For Each myShape In ActiveSheet.Shapes
            ' 図が写真なら
            If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
                ' 写真の名前を取得
                PictureName(j) = myShape.Name
                ' 写真の枚数をカウント
                j = j + 1
            End If
        Next
        ' 全ての図に対して
        For k = 0 To j - 1
            ' 撮影日データを取得
            myDate = Range(PictureDateBuffer & Format(MinDataBuffer + pageNo(ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Row, _
                ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Column))).Value
            ' 図のセルを選択
            Cells(ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Row, ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Column).Select
            ' 撮影日の表示形式フラグをエンコード
            myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
            ' 撮影日のテキストボックスを作図
            Call AddPictureDate(PictureName(k), PictureDateFormat(myDate, myType))
        Next k
    End If
    ' 撮影日表示フラグをセット
    Range(PictureDateFlag).Value = 1
    Erase PictureName
    ' 現在のセルを選択
    Cells(CurrentRow, CurrentColumn).Select
End Sub

Sub AddPictureDate(ByVal myPictureName As String, ByVal myDate As String)
' 撮影日のテキストボックスを作図
    Dim WeekLen As Integer
    ' 曜日表示する場合の文字数補正
    If Range(WeekDisp).Value <> 0 And Range(WeekLang).Value = 0 Then
        WeekLen = 2
    Else
        WeekLen = 1
    End If
    ' テキストボックスを作図
    With ActiveSheet.Shapes.AddTextbox _
        (msoTextOrientationHorizontal, _
        ActiveCell.Left, ActiveCell.Top, ((Len(myDate) + WeekLen) * Range(DateFontSize).Value / 2), (Range(DateFontSize).Value + DateHeightOffset))
        ' テキストボックスの名前を設定
        .Name = PictureDateName(myPictureName)
        ' テキスト文字
        .TextFrame2.TextRange.Characters.Text = myDate
        ' フォントサイズ
        .TextFrame2.TextRange.Characters.Font.Size = Range(DateFontSize).Value
        ' フォントの色
        .TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = _
            RGB(Range(DateFontColorR).Value, Range(DateFontColorG).Value, Range(DateFontColorB).Value)
        ' 太字にする
        .TextFrame2.TextRange.Characters.Font.Bold = (Range(DateFontBold).Value <> 0)
        ' 文字の左マージン
        .TextFrame2.MarginLeft = 0
        ' 文字の右マージン
        .TextFrame2.MarginRight = 0
        ' 文字の上マージン
        .TextFrame2.MarginTop = 0
        ' 文字の下マージン
        .TextFrame2.MarginBottom = 0
        ' 文字の折り返しなし
        .TextFrame2.WordWrap = msoFalse
        ' テキストボックスの文字を右詰め
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        ' テキストボックスの文字の上下を中央ぞろえ
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        ' 線無し
        .Line.Visible = False
        ' 塗りつぶし無し
        .Fill.Visible = False
        ' 位置決め
        .Top = Selection.Top
        .Left = Selection.Left
        ' テキストボックスに、マクロを登録
        .OnAction = "PictureDateClick"
    End With
    ' テキストボックスの位置決め
    Call PictureDatePosition(myPictureName)
End Sub

Sub PictureDatePosition(ByVal myPictureName As String)
' 撮影日のテキストボックスの位置決め
    Dim PictureWidth As Double, PictureHeight As Double
    '写真の幅、高さを取得
    PictureWidth = ActiveSheet.Shapes(myPictureName).Width
    PictureHeight = ActiveSheet.Shapes(myPictureName).Height
    ' 写真のセルを選択
    ActiveSheet.Shapes(myPictureName).TopLeftCell.MergeArea.Select
    ' 写真が９０°または２７０°回転している場合
    If ((ActiveSheet.Shapes(myPictureName).Rotation = 90 Or ActiveSheet.Shapes(myPictureName).Rotation = 270)) Then
        ' 撮影日を写真の右下に位置決め
        With ActiveSheet.Shapes(PictureDateName(myPictureName))
            .Top = Selection.Top + (Selection.Height - PictureWidth) / 2 + PictureWidth - .Height - Range(DateYOffset).Value
            .Left = Selection.Left + (Selection.Width - PictureHeight) / 2 + PictureHeight - .Width - Range(DateXOffset).Value
        End With
    ' 写真が０°または１８０°回転している場合
    Else
        ' 撮影日を写真の右下に位置決め
        With ActiveSheet.Shapes(PictureDateName(myPictureName))
            .Top = Selection.Top + (Selection.Height - PictureHeight) / 2 + PictureHeight - .Height - Range(DateYOffset).Value
            .Left = Selection.Left + (Selection.Width - PictureWidth) / 2 + PictureWidth - .Width - Range(DateXOffset).Value
        End With
    End If
End Sub

Function PictureDateFormat(ByVal myDate As String, ByVal myType As Integer) As String
' 撮影日のフォーマット設定
    If myDate <> "" Then
        ' 半角文字に変換
        myDate = StrConv(myDate, vbNarrow)
        ' 表示できない文字を削除
        myDate = Replace(myDate, "?", "")
        ' 曜日表示「（」がある場合、年月日と時刻を抽出
        If (InStr(myDate, "(") > 8) And (InStr(myDate, "(") <= 11) Then
            myDate = Left(myDate, InStr(myDate, "(") - 1) & " " & Right(myDate, Len(myDate) - InStrRev(myDate, " "))
        End If
    End If
    ' 撮影日表示形式が日付だけの場合（myTypeの値をデコードして比較）
    If (myType Mod 2) <> 0 Then
        ' 引数が日時データなら
        If IsDate(myDate) Then
            ' 表示形式を日付のみにする
            myDate = Format(myDate, "yyyy/mm/dd")
            ' 曜日を表示する場合（myTypeの値をデコードして比較）
            If ((myType \ 4) Mod 2) <> 0 Then
                ' 曜日を英語で表示する場合（myTypeの値をデコードして比較）
                If ((myType \ 8) Mod 2) <> 0 Then
                    myDate = Format(myDate, "yyyy/mm/dd(ddd)")
                ' 曜日を日本語で表示
                Else
                    myDate = Format(myDate, "yyyy/mm/dd(aaa)")
                End If
            End If
        ' 引数myDateが日時テータとして扱えない場合
        Else
            ' 日時を＊で表示
            myDate = "****/**/**"
        End If
    ' 撮影日表示形式が日付と時刻の場合
    Else
        ' 引数が日時データなら
        If IsDate(myDate) Then
            ' 表示形式を日付と時刻にする
            myDate = Format(myDate, "yyyy/mm/dd h:nn")
            ' 曜日を表示する場合（myTypeの値をデコードして比較）
            If ((myType \ 4) Mod 2) <> 0 Then
                ' 曜日を英語で表示する場合（myTypeの値をデコードして比較）
                If ((myType \ 8) Mod 2) <> 0 Then
                    myDate = Format(myDate, "yyyy/mm/dd(ddd) h:nn")
                ' 曜日を日本語で表示
                Else
                    myDate = Format(myDate, "yyyy/mm/dd(aaa) h:nn")
                End If
            End If
        ' 引数myDateが日時データとして扱えない場合
        Else
            ' 日付と時刻を＊で表示
            myDate = "****/**/** **:**"
        End If
    End If
    ' 撮影日の日付区切り記号に「．」を使う場合（myTypeの値をデコードして比較）
    If ((myType \ 2) Mod 2) <> 0 Then
        ' 日付の「/」を「.」に置き換える
        myDate = Replace(myDate, "/", ".")
    End If
    ' 変換した日付時刻を戻り値にする
    PictureDateFormat = myDate
End Function

Function PictureDateName(ByVal myPictureName As String) As String
' 撮影日テキストボックスの名前を生成
    ' 写真の名前の最後の１文字を「DateSuffix」で定義している文字に置き換える
    PictureDateName = Left(myPictureName, (Len(myPictureName) - 1)) & DateSuffix
End Function

Sub PictureDateClick()
' 撮影日テキストボックスをクリックした場合
    ' 現在のシートを選択
    ActiveSheet.Select
    ' ユーザーフォーム９を呼び出す
    UserForm9.Show vbModal
End Sub

Function GetPictureNoFromDate(ByVal myName As String) As Integer
' 撮影日テキストボックスの名前から写真のページ番号を求める
    Dim i As Integer, j As Integer, myShape As Shape, PictureName() As String, myPictureDate As String
    ' 写真の枚数を求める
    i = 0
    For Each myShape In ActiveSheet.Shapes
        If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
            i = i + 1
        End If
    Next
    ' 写真が１枚以上なら
    If i > 0 Then
        ' 動的配列の宣言
        ReDim PictureName(i - 1)
        j = 0
        ' すべての図に対して
        For Each myShape In ActiveSheet.Shapes
            ' 図が写真なら
            If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
                ' 写真の名前を取得
                PictureName(j) = myShape.Name
                ' 写真の枚数をカウント
                j = j + 1
            End If
        Next
        ' すべての写真の名前に対して
        For j = 0 To i - 1
            ' 撮影日テキストボックスの名前が引数と一致すれば
            If PictureDateName(PictureName(j)) = myName Then
                ' 繰り返しを抜ける
                Exit For
            End If
        Next j
        ' 写真のページ番号を戻り値に設定
        GetPictureNoFromDate = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
            ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
    ' 写真が無い場合は
    Else
        ' 戻り値をゼロにしておく
        GetPictureNoFromDate = 0
    End If
    ' 配列変数の解放
    Erase PictureName
End Function

Sub PictureDateDispSequence()
    Dim Ans As Integer
    ' 撮影日時表示フラグがＯＮなら
    If Range(PictureDateFlag).Value <> 0 Then
        Ans = MsgBox("撮影日時を消去しますか？" & vbCrLf & _
            "消去する場合は「はい」を" & vbCrLf & _
            "撮影日時の表示を変更する場合は「いいえ」をクリックしてください。", vbYesNoCancel + vbInformation, "確認")
        If Ans = vbYes Then
            ' 撮影日時消去処理を呼び出す
            Call PictureDateOFF
        ElseIf Ans = vbNo Then
            ' 撮影日時の表示選択
            UserForm8.Show vbModal
        End If
    ' 撮影日時表示フラグがＯＦＦなら
    Else
        ' 撮影日時の表示選択
        UserForm8.Show vbModal
    End If
End Sub

Sub DeleteLastPages()
' 末尾のページ削除処理
    Dim myBottomCount As Long
    Dim myPictureNo As Integer, myMaxNo As Integer
    Dim myPicture As Shape
    Application.ScreenUpdating = False
    Application.StatusBar = "末尾のページを削除しています。お待ちください。"
    ' 写真の最大通し番号
    myMaxNo = 1
    ' すべての図にたいして
    For Each myPicture In ActiveSheet.Shapes
        myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' 図の行数がページの最大値より大きい場合
        If myPicture.BottomRightCell.Row > MaxPageRow Then
            ' メッセージを表示する
            MsgBox "写真が最大ページを超えて貼り付けられています。", vbOKOnly + vbExclamation, "お知らせ"
            ' 処理を終了
            Application.ScreenUpdating = True
            Application.StatusBar = False
            Exit Sub
        ElseIf (myPictureNo > myMaxNo) Then
            ' 図の最大通し番号を求める
            myMaxNo = myPictureNo
        End If
    Next
    ' 最終行を計算
    myBottomCount = (PictureRow(myMaxNo) \ 33) * 33 + 33
    ' ページを削除
    If myBottomCount < 65536 Then
        Range("A" & Format(myBottomCount + 1) & ":G65536").EntireRow.Delete
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ' 写真の最大ページを選択
    Range(PictureColumn(myMaxNo) & Format(PictureRow(myMaxNo))).Select
End Sub
