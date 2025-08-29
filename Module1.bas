Attribute VB_Name = "Module1"
Option Explicit
Const myMinRow As Long = 2
Const MaxPageRow As Long = 65472
Const MaxPageNo As Integer = 5952
Const IndexRowHeight As Double = 17.25 ' �ԍ��s�̃Z���̍���
Const CommentRowHeight As Double = 30.75 ' �s�̃Z���̍���
Public Const TopRowHeight As Double = 15# ' �P�s�ڂ̃Z���̍���
Public Const TempRowHeight As Double = 300# ' �P�s�ڂŉ摜���Y����΍��p�Z���̍���
Public ShortCutFlag As Boolean ' �V���[�g�J�b�g�\���p�t���O
Public SwapSourceNo As Integer, SwapDestNo As Integer ' �ʐ^�̓���ւ����A����ւ���̃y�[�W�ԍ�
Public SwapSourceName As String, SwapDestName As String ' �ʐ^�̓���ւ����A����ւ���̎ʐ^��
Public SwapSourceSheet As String, SwapDestSheet As String ' �ʐ^�̓���ւ����A����ւ���̃V�[�g��
Public ComboList1 As Long ' �R���{�{�b�N�X�P�̃��X�g�ԍ��ۑ��p�i���[�U�[�t�H�[���S�Ŏg�p�j
Public ComboList2 As Long ' �R���{�{�b�N�X�Q�̃��X�g�ԍ��ۑ��p�i���[�U�[�t�H�[���S�Ŏg�p�j
Public ComboList3 As Long ' �R���{�{�b�N�X�R�̃��X�g�ԍ��ۑ��p�i���[�U�[�t�H�[���S�Ŏg�p�j
Public ComboList4 As Long ' �R���{�{�b�N�X�S�̃��X�g�ԍ��ۑ��p�i���[�U�[�t�H�[���S�Ŏg�p�j
Public ComboList5 As Long ' �R���{�{�b�N�X�T�̃��X�g�ԍ��ۑ��p�i���[�U�[�t�H�[���S�Ŏg�p�j
Public ComboList6 As Long ' �R���{�{�b�N�X�U�̃��X�g�ԍ��ۑ��p�i���[�U�[�t�H�[���S�Ŏg�p�j
Public Const PictureNameBuffer As String = "K" ' �ʐ^�̃t�@�C�����ۑ��p�Z���̗�ԍ�
Public Const PictureDateBuffer As String = "L" ' �ʐ^�̎B�e���ۑ��p�Z���̗�ԍ�
Public Const MinDataBuffer As Integer = 12 ' �ʐ^�̎B�e���A�t�@�C�����ۑ��p�Z���̊J�n�s�ԍ��I�t�Z�b�g
Public Const CutDataBuffer As Integer = 12 ' �؂������ʐ^�̎B�e���A�t�@�C�����ۑ��p�Z���̍s�ԍ�
Public Const PictureDateFlag As String = "K2" ' �ʐ^�̎B�e���̕\�������邩�ǂ����̃t���O�ۑ��p�Z���Ԓn
Public Const PictureDateType As String = "L2" ' �B�e���\���^�C�v�̃t���O�ۑ��p�Z���Ԓn
Public Const DateSuffix As String = "D" ' �B�e���e�L�X�g�{�b�N�X�̖��O���ʗp����
Public CutDataSheet As String ' �ʐ^�̐؂�����s�����V�[�g��
Public Const DateFontSize As String = "K3" ' �B�e���̃t�H���g�T�C�Y�ۑ��p�̃Z���Ԓn
Public Const DateFontBold As String = "K4" ' �B�e���̃t�H���g�����t���O�ۑ��p�̃Z���Ԓn
Public Const DateFontColorR As String = "L3" ' �B�e���̃t�H���g�F�i�ԁj�ۑ��p�̃Z���Ԓn
Public Const DateFontColorG As String = "L4" ' �B�e���̃t�H���g�F�i�΁j�ۑ��p�̃Z���Ԓn
Public Const DateFontColorB As String = "L5" ' �B�e���̃t�H���g�F�i�j�ۑ��p�̃Z���Ԓn
Public Const DateHeightOffset As Double = 0# ' �B�e���e�L�X�g�{�b�N�X�̍����}�[�W��
Public Const DateXOffset As String = "K5" ' �B�e���e�L�X�g�{�b�N�X�̉E����̃I�t�Z�b�g�ۑ��p�Z���Ԓn
Public Const DateYOffset As String = "K6" ' �B�e���e�L�X�g�{�b�N�X�̉�����̃I�t�Z�b�g�ۑ��p�Z���Ԓn
Public Const XUnit As Double = 0.33 ' �E����̃I�t�Z�b�g�̒P�ʕϊ��l
Public Const YUnit As Double = 0.325 ' ������̃I�t�Z�b�g�̒P�ʕϊ��l
Public Const DateSeparator As String = "L6" ' �B�e���̋�؂���u�D�v�ɂ���t���O�ۑ��p�̃Z���Ԓn
Public Const WeekDisp As String = "K7" ' �B�e���ɗj����\������t���O�ۑ��p�̃Z���Ԓn
Public Const WeekLang As String = "K8" ' �B�e���̗j���̌���t���O�ۑ��p�̃Z���Ԓn
Public SelectedOrder() As Integer ' �摜�̏��ԁi��FImage1��2���ڂȂ� SelectedOrder(1) = 2�j
Public SelectedPaths() As String  ' �I�����ꂽ�摜�̃p�X
Public PasteStartCell As Range ' �� �W�����W���[���ɐ錾

Function AddPages(myInsertCount As Integer) As Integer
' ������ǉ�����֐��i�����͒ǉ����閇���j
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim myPageCount As Integer
    Dim myRange As Range
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �߂�l���|�P�ɂ���
        AddPages = -1
        ' �������I��
        Exit Function
    End If
    ' �y�[�W�J�E���^
    j = 1
    ' �y�[�W�}������
    k = 0
    ' �ŏI�y�[�W
    m = 1
    ' �擪�̃y�[�W�̃Z����I������
    Set myRange = Range(PictureColumn(1) & Format(PictureRow(1)))
    myRange.Select
    ' �X�e�[�^�X�o�[�ɏ�Ԃ�\������
    Application.StatusBar = "�y�[�W��ǉ����Ă��܂��B���҂����������B"
    ' �J��Ԃ�����
    Do
        ' �Z�����P�O��������Ă���ꍇ
        If (myRange.MergeArea.Rows.Count = 10) And (myRange.MergeArea.Columns.Count = 1) Then
            ' �y�[�W�J�E���^�����Z
            j = j + 1
            ' �ŏI�y�[�W���L��
            m = j
            ' ���̃y�[�W�̃Z����I��
            Set myRange = Range(PictureColumn(j) & Format(PictureRow(j)))
            myRange.Select
            ' �ǉ�����y�[�W���ő�l�𒴂����ꍇ
            If j > MaxPageNo Then
                ' �߂�l���|�Q�ɂ���
                AddPages = -2
                Range(PictureColumn(MaxPageNo) & Format(PictureRow(MaxPageNo))).Select
                Application.StatusBar = False
                ' �������I��
                Exit Function
            End If
        ' �Z������������Ă��Ȃ��ꍇ�܂��͊���̌����łȂ��ꍇ
        Else
            ' �Z������������Ă���ꍇ
            If myRange.MergeCells Then
                ' �Z���̌���������
                myRange.MergeCells = False
            End If
            ' �s�̍�����ݒ肷��
            myRange.Offset(-1, 0).RowHeight = TopRowHeight
            ' �ǉ�����s�����v�Z����
            myPageCount = (3 - ((j - 1) Mod 3))
            ' ������ǉ�
            For i = 1 To myPageCount
                ' �Z��������
                Range(PictureColumn(j) & Format(PictureRow(j)) & _
                    ":" & PictureColumn(j) & Format(PictureRow(j) + 9)).Merge
                ' �]��������}������
                With Range(PictureColumn(j) & Format(PictureRow(j))).MergeArea
                    ' �Z���̒l
                    .Value = "�]��"
                    ' �t�H���g�T�C�Y
                    .Font.Size = 72
                    ' �����̐F
                    .Font.Color = RGB(192, 192, 192)
                    ' �����ʒu�𒆉����킹
                    .HorizontalAlignment = xlCenter
                    ' �����ʒu�𒆉����킹
                    .VerticalAlignment = xlCenter
                End With
                ' �s�̍�����ݒ肷��
                Range("A" & Format(PictureRow(j)) & ":A" & Format(PictureRow(j) + 1)).RowHeight = IndexRowHeight
                ' �r��������
                With Range(CommentColumn(j) & Format(PictureRow(j)))
                    ' ����������
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    ' �r���̐F
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                    ' �y�[�W�ԍ���}������
                    .Value = "No." & Format(j)
                    ' �����ʒu�����l��
                    .HorizontalAlignment = xlLeft
                    ' �����ʒu�𒆉����킹
                    .VerticalAlignment = xlCenter
                    ' �t�H���g�T�C�Y
                    .Font.Size = 11
                    ' �����̐F
                    .Font.Color = RGB(0, 0, 0)
                End With
                With Range(CommentColumn(j) & Format(PictureRow(j) + 2) & _
                    ":" & CommentColumn(j) & Format(PictureRow(j) + 8))
                    ' �I��͈͂̉��ɓ_��������
                    .Borders(xlEdgeBottom).LineStyle = xlDot
                    ' �r���̐F
                    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
                    ' �I��͈͂̒��ɓ_��������
                    .Borders(xlInsideHorizontal).LineStyle = xlDot
                    ' �r���̐F
                    .Borders(xlInsideHorizontal).Color = RGB(0, 0, 0)
                    ' �����ʒu
                    .HorizontalAlignment = xlLeft
                    ' �����ʒu
                    .VerticalAlignment = xlCenter
                    ' �t�H���g�T�C�Y
                    .Font.Size = 11
                    ' �����̐F
                    .Font.Color = RGB(0, 0, 0)
                    ' �k�����đS�̂�\��
                    .WrapText = False
                    .ShrinkToFit = True
                    ' �s�̍�����ݒ肷��
                    .RowHeight = CommentRowHeight
                End With
                ' �s�̍�����ݒ肷��
                Range("A" & Format(PictureRow(j) + 9)).RowHeight = CommentRowHeight
                Range("A" & Format(PictureRow(j) + 10)).RowHeight = TopRowHeight
                ' ���̃y�[�W��
                j = j + 1
                ' ���̃y�[�W�̃Z����I��
                Set myRange = Range(PictureColumn(j) & Format(PictureRow(j)))
                myRange.Select
            Next i
            ' �y�[�W�}�����������Z
            k = k + 1
            ' �y�[�W�}�������������ƈ�v�����ꍇ
            If k >= myInsertCount Then
                ' �J��Ԃ��������I���
                Exit Do
            End If
            ' �ǉ�����y�[�W���ő�l�𒴂����ꍇ
            If j > MaxPageNo Then
                ' �߂�l���|�Q�ɂ���
                AddPages = -2
                Range(PictureColumn(MaxPageNo) & Format(PictureRow(MaxPageNo))).Select
                Application.StatusBar = False
                ' �������I��
                Exit Function
            End If
        End If
    Loop
    ' �ǉ����������̐擪�̃Z����I������
    Range(PictureColumn(m) & Format(PictureRow(m))).MergeArea.Select
    ' �߂�l���[���ɂ���i����I���j
    AddPages = 0
    Application.StatusBar = False
End Function

Sub PictureRotation()
' �ʐ^�����N���b�N�����ꍇ�̏���
    ' ���݂̃V�[�g��I��
    ActiveSheet.Select
    ' �ʐ^�̈ړ����y�[�W�ԍ����Z�b�g����Ă��Ȃ����
    If SwapSourceNo = 0 Then
        ' ���[�U�[�t�H�[���R���Ăяo��
        UserForm3.Show vbModal
    ' �ʐ^�̈ړ����y�[�W�ԍ����Z�b�g����Ă����
    Else
        ' �ʐ^�̓���ւ���y�[�W�ԍ����擾
        SwapDestNo = pageNo(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row, ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column)
        ' �ʐ^�̓���ւ���ʐ^�����擾
        SwapDestName = ActiveSheet.Shapes(Application.Caller).Name
        ' �ʐ^�̓���ւ���V�[�g�����擾
        SwapDestSheet = ActiveSheet.Name
        ' �ʐ^�̓���ւ����y�[�W�ԍ��Ǝʐ^�̓���ւ���y�[�W�ԍ����قȂ��
        If (SwapSourceNo <> SwapDestNo) Or (SwapSourceSheet <> SwapDestSheet) Then
            ' ���[�U�[�t�H�[���T���\���ɂ�
            UserForm5.Hide
            ' ���[�T�[�t�H�[���U���Ăяo��
            UserForm6.Show vbModal
        ' �ʐ^�̓���ւ����y�[�W�ԍ��Ǝʐ^�̓���ւ���y�[�W�ԍ��������Ȃ��
        Else
            ' ���b�Z�[�W��\��
            MsgBox "�ʐ^�̓���ւ����Ǝʐ^�̓���ւ��悪�����ł��B", vbOKOnly + vbExclamation, "���m�点"
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

        ' �Z���T�C�Y�擾
        cellWidth = ActiveCell.MergeArea.Width - 4
        cellHeight = ActiveCell.MergeArea.Height - 4

        ' ���T�C�Y�ɖ߂�
        .ScaleHeight 1, msoTrue
        .ScaleWidth 1, msoTrue

        ' �Z���Ɏ��܂�悤�ɏc������ێ����ďk��
        If .Width > cellWidth Or .Height > cellHeight Then
            scaleRatio = Application.Min(cellWidth / .Width, cellHeight / .Height)
            .ScaleWidth scaleRatio, msoTrue
            .ScaleHeight scaleRatio, msoTrue
        End If

        ' �����ɔz�u
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
' �ʐ^�̉�]�����i�����͐}�`�̖��O�A��]�p�x�j
    Dim myWidth As Double, myHeight As Double, myAspectRatio As Double
    Application.ScreenUpdating = False
    With ActiveSheet.Shapes(myShapeName)
        .TopLeftCell.MergeArea.Select
        ' �u�`�Q�v�Z���Ŏʐ^���c���łX�O���܂��͂Q�V�O���̏ꍇ�Ɉʒu���߂������΍�
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TempRowHeight
        End If
        myWidth = .Width
        myHeight = .Height
        ' �ʐ^�̏c������v�Z
        myAspectRatio = myWidth / myHeight
        ' �c����̌Œ����������
        .LockAspectRatio = msoFalse
        ' �ʐ^�𐳕��`�ɂ���
        If myWidth > myHeight Then
            .Height = myWidth
        ElseIf myWidth < myHeight Then
            .Width = myHeight
        End If
        ' �ʐ^����]������
        .Rotation = .Rotation + myDegree
        ' �c����𕜌�����
        If .Rotation = 90 Or .Rotation = 270 Then
        ' ��]�p�x���X�O���܂��͂Q�V�O���̏ꍇ
            If myWidth > myHeight Then
            ' �ʐ^�̕����������傫���ꍇ
                ' �ʐ^�̕���g�ɍ��킹��
                .Width = Int(Selection.Height) - 4
                ' �ʐ^�̍����𕜌�
                .Height = .Width / myAspectRatio
            ElseIf myWidth < myHeight Then
            ' �ʐ^�̍����������傫���ꍇ
                ' �ʐ^�̍�����g�ɍ��킹��
                .Height = Int(Selection.Width) - 4
                ' �ʐ^�̕��𕜌�
                .Width = .Height * myAspectRatio
            End If
            ' �c������Œ肷��
            .LockAspectRatio = msoTrue
            ' �ʐ^�̕��܂��͍������Z�����傫���Ȃ�΃Z���ɍ��킹��
            If .Width > Int(Selection.Height) - 4 Then
                .Width = Int(Selection.Height) - 4
            ElseIf .Height > Int(Selection.Width) - 4 Then
                .Height = Int(Selection.Width) - 4
            End If
        Else
        ' ��]�p�x���O���܂��͂P�W�O���̏ꍇ
            If myWidth > myHeight Then
            ' �ʐ^�̕����������傫���ꍇ
                ' �ʐ^�̕���g�ɍ��킹��
                .Width = Int(Selection.Width) - 4
                ' �ʐ^�̍����𕜌�
                .Height = .Width / myAspectRatio
            ElseIf myWidth < myHeight Then
            ' �ʐ^�̍����������傫���ꍇ
                ' �ʐ^�̍�����g�ɍ��킹��
                .Height = Int(Selection.Height) - 4
                ' �ʐ^�̕��𕜌�
                .Width = .Height * myAspectRatio
            End If
            ' �c������Œ肷��
            .LockAspectRatio = msoTrue
            ' �ʐ^�̕��܂��͍������Z�����傫���Ȃ�΃Z���ɍ��킹��
            If .Height > Int(Selection.Height) - 4 Then
                .Height = Int(Selection.Height) - 4
            ElseIf .Width > Int(Selection.Width) - 4 Then
                .Width = Int(Selection.Width) - 4
            End If
        End If
        ' �ʐ^�̈ʒu����
        .Top = Selection.Top + ((Selection.Height - .Height) / 2)
        .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        ' �u�`�Q�v�Z���Ŏʐ^���c���łX�O���܂��͂Q�V�O���̏ꍇ�Ɉʒu���߂������΍�
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TopRowHeight
        End If
    End With
    Application.ScreenUpdating = True
End Sub

Sub PreviewPrint()
' ����v���r���[��\������
    Dim myTopCount As Long, myBottomCount As Long
    Dim myPictureNo As Integer, myMaxNo As Integer
    Dim myPicture As Shape
    ' �ʐ^�̍ő�y�[�W�ԍ�
    myMaxNo = 1
    ' ���ׂĂ̐}�ɂ�������
    For Each myPicture In ActiveSheet.Shapes
        myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' �}�̍s�����y�[�W�̍ő�l���傫���ꍇ
        If myPicture.BottomRightCell.Row > MaxPageRow Then
            ' ���b�Z�[�W��\������
            MsgBox "�ʐ^���ő�y�[�W�𒴂��ē\��t�����Ă��܂��B", vbOKOnly + vbExclamation, "���m�点"
            ' �������I��
            Exit Sub
        ElseIf (myPictureNo > myMaxNo) Then
            ' �}�̍ő�s�������߂�
            myMaxNo = myPictureNo
        End If
    Next
    myTopCount = 1
    myBottomCount = (PictureRow(myMaxNo) \ 33) * 33 + 33
    With ActiveSheet
        ' ����͈͂�ݒ肷��
        .PageSetup.PrintArea = "A" & Format(myTopCount) & ":B" & Format(myBottomCount)
        ' ����̕������c�ɐݒ肷��
        .PageSetup.Order = xlDownThenOver
        ' ����v���r���[��\������
        .PrintPreview
        ' ����͈͂���������
        .PageSetup.PrintArea = False
    End With
End Sub

Sub AddPageProc(ByVal myPage As Integer)
' �y�[�W�̒ǉ������i�����͒ǉ������j
    ' ��ʕ\���̍X�V�����Ȃ��悤�ɂ���
    Application.ScreenUpdating = False
    ' �����̒ǉ��֐��i�ǉ������j
    Select Case AddPages(myPage)
        ' �߂�l�ɂ�鏈��
        Case -1
            MsgBox "�V�[�g���ی삳��Ă��܂��B" _
                & vbCrLf & "�ی���������Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Case -2
            MsgBox "����ȏ�y�[�W��ǉ��ł��܂���B", vbOKOnly + vbExclamation, "���m�点"
    End Select
    ' ��ʕ\���̍X�V��������
    Application.ScreenUpdating = True
    ActiveCell.Select
End Sub

Sub GetMultiPicture()
' �ꊇ��荞�ݏ���
    Dim myPicture As Shape
    Dim myRange As Range
    Dim myPictureNo As Integer, myMaxNo As Integer
    ' �}�`�̍ő�y�[�W��
    myMaxNo = 0
    ' ���ׂĂ̐}�`�ɑ΂���
    For Each myPicture In ActiveSheet.Shapes
        ' �}�`�̍���Z�����ŏ��Z���ȏ�̏ꍇ
        If myPicture.TopLeftCell.Row >= myMinRow Then
            ' �}�`�̃y�[�W�������߂�
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' �y�[�W�̍ő吔�����߂�
            If myMaxNo < myPictureNo Then
                myMaxNo = myPictureNo
            End If
        End If
    Next
    myMaxNo = myMaxNo + 1
    ' �ő�y�[�W�𒴂���ꍇ�͏����𒆒f
    If myMaxNo > MaxPageNo Then
        MsgBox "�ʐ^���ő�y�[�W�𒴂��܂��B" _
            & vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
    Else
        ' �ʐ^�̃y�[�W����Z����ݒ�
        Set myRange = Range(PictureColumn(myMaxNo) & Format(PictureRow(myMaxNo)))
        ' �Z����I��
        myRange.MergeArea.Select
        ' �����Ɉꊇ�捞���Ăяo��
        Call GetMultiPictureFromHere
    End If
End Sub

Sub GetMultiPictureFromHere()
' �����Ɉꊇ��荞�ݏ���
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
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        MsgBox "�V�[�g���ی삳��Ă��܂��B" _
            & vbCrLf & "�ی���������Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        ' �������I��
        Exit Sub
    End If
    
    Dim targetCell As Range
    Set targetCell = ActiveCell ' ��F���ݑI������Ă���Z����Ώۂɂ���
    Set PasteStartCell = targetCell
    
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    StartNo = CurrentNo
    ' ���݂̃y�[�W�ԍ����u�O�v�ȉ��̏ꍇ�������I��
    If CurrentNo <= 0 Then
        ' �������I��
        Exit Sub
    End If
    ' �t�@�C���_�C�A���O�{�b�N�X���J��
    
    
    
   With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = True
    .Title = "�摜��I�����Ă�������"
    .ButtonName = "��荞��"
    .Filters.Clear
    .Filters.Add "�摜", "*.JPG;*.JPEG;*.BMP;*.TIF;*.TIFF;*.PNG;*.GIF;*.HEIC", 1

    If .Show = -1 Then
        mySelectedItemsCount = .SelectedItems.Count

        ReDim SelectedPaths(1 To mySelectedItemsCount)
        ReDim SelectedOrder(1 To mySelectedItemsCount)
        
        ' �t�@�C���p�X�� SelectedPaths �Ɋi�[
    For i = 1 To mySelectedItemsCount
        SelectedPaths(i) = .SelectedItems(i)
    Next i



            ' ?? �����ɏ�����������ǉ��I
    For i = 1 To 36
        With UserForm10.Controls("Image" & i)
            Set .Picture = Nothing
        End With
        UserForm10.Controls("Label" & i).Caption = ""
    Next i

    ' �摜�\�������iImage1�`ImageN�j
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

    ' �t�H�[���\��
    UserForm10.Show
End If


End With

    Application.ScreenUpdating = False
    ' �������O�̎ʐ^�ƒ����t�@�C�����̎ʐ^���J�E���g����
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
    ' ��荞�ݖ������v�Z
    myFnameCount = mySelectedItemsCount - myFailureCount - myLongNameCount
    ' ��荞�ݖ������P���ȏ�Ȃ�
    If myFnameCount > 0 Then
        ReDim myFname(myFnameCount - 1)
        j = 0
        ' �I���t�@�C���������肩����
        For i = 0 To mySelectedItemsCount - 1
            mySamePictureExist = False
            myPictureName = Right(myFilename(i), Len(myFilename(i)) - InStrRev(myFilename(i), "\"))
            ' �������O�̎ʐ^���m�F
            For Each myPicture In ActiveSheet.Shapes
                If myPicture.Name = myPictureName Then
                    mySamePictureExist = True
                    Exit For
                End If
            Next
            ' �ʐ^�̃t�@�C�������R�O���ȉ��Ŏʐ^���������O�łȂ����
            If mySamePictureExist = False And Len(myPictureName) <= 30 Then
                ' ��荞�ݎʐ^��z��ɑ��
                myFname(j) = myFilename(i)
                j = j + 1
            End If
        Next i
        ' �ʐ^�̍ő�y�[�W�Ɩ��������߂�
        MaxNo = 0
        m = 0
        ' ���ׂĂ̐}�ɂ�������
        For Each myPicture In ActiveSheet.Shapes
            ' �}�̃y�[�W�����߂�
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' ���݂̃Z�����}�̃y�[�W���傫���ꍇ
            If CurrentNo <= myPictureNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' �������J�E���g
                m = m + 1
                ' �}�̃y�[�W���ő�y�[�W���傫���ꍇ
                If MaxNo < myPictureNo Then
                    ' �ő�y�[�W���X�V
                    MaxNo = myPictureNo
                End If
            End If
        Next
        ' �ʐ^�̍ŏ��y�[�W�����߂�
        MinNo = MaxNo
        ' ���ׂĂ̐}�ɂ�������
        For Each myPicture In ActiveSheet.Shapes
            ' �}�̃y�[�W�����߂�
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' ���݂̃Z�����}�̃y�[�W���傫���ꍇ
            If CurrentNo <= myPictureNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' �}�̃y�[�W���ŏ��y�[�W��菬�����ꍇ
                If MinNo > myPictureNo Then
                    ' �ŏ��y�[�W���X�V
                    MinNo = myPictureNo
                End If
            End If
        Next
        ' ���݃Z���ȍ~�Ɏʐ^���Ȃ��ꍇ
        If MaxNo = 0 Then
            ' �ő�y�[�W�����݂̃y�[�W�ԍ��ɐݒ�
            MaxNo = CurrentNo
            ' ��������Ă���Z���̃J�E���^
            j = 0
            ' ���肩��������
            Do
                ' �ő�y�[�W�̃Z����I��
                Set myRange = Range(PictureColumn(MaxNo) & Format(PictureRow(MaxNo)))
                myRange.MergeArea.Activate
                ' �Z������������Ă���ꍇ
                If (ActiveCell.MergeArea.Rows.Count = 10) And (ActiveCell.MergeArea.Columns.Count = 1) Then
                    ' ��������Ă���Z���̃J�E���^�����Z
                    j = j + 1
                    ' �ő�y�[�W�����Z
                    MaxNo = MaxNo + 1
                    ' �ő�y�[�W�𒴂���ꍇ
                    If (j < myFnameCount) And (MaxNo > MaxPageNo) Then
                        MsgBox "�ʐ^���ő�y�[�W�𒴂��܂��B" & _
                            vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                ' �Z������������Ă��Ȃ��ꍇ
                Else
                    ' �ǉ�����y�[�W�����v�Z
                    PageInsertCount = ((myFnameCount - j - 1) \ 6) + 1
                    ' �ǉ�����y�[�W���ő�l�𒴂���ꍇ
                    If (pageNo(ActiveCell.MergeArea.Row, ActiveCell.MergeArea.Column) - 1 + PageInsertCount * 6) _
                        > MaxPageNo Then
                        MsgBox "�ʐ^���ő�y�[�W�𒴂��܂��B" & _
                            vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                    ' �y�[�W��ǉ�����
                    If AddPages(PageInsertCount) < 0 Then
                        MsgBox "�y�[�W��ǉ��ł��܂���B" & _
                            vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                End If
            ' �Z���̌�������荞�ݖ������Ȃ��Ԃ��肩����
            Loop While j < myFnameCount
        ' ���݃Z���ȍ~�Ɏʐ^������ꍇ
        Else
            ' �]���Z���̐�
            BlankCount = MinNo - CurrentNo
            ' �]���Z���̐�����荞�ݖ�����菭�Ȃ��ꍇ
            If BlankCount < myFnameCount Then
                ' ��������Ă���Z���̃J�E���^
                j = 0
                ' �ő�y�[�W�����̃R�}��
                MaxNo = MaxNo + 1
                ' ���肩��������
                Do
                    ' �ő�y�[�W�𒴂���ꍇ
                    If (j < (myFnameCount - BlankCount)) And (MaxNo > MaxPageNo) Then
                        MsgBox "�ʐ^���ő�y�[�W�𒴂��܂��B" & _
                            vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
                        Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                        Exit Sub
                    End If
                   ' �ő�y�[�W�̃Z����I��
                    Set myRange = Range(PictureColumn(MaxNo) & Format(PictureRow(MaxNo)))
                    myRange.MergeArea.Activate
                    ' �Z������������Ă���ꍇ
                    If (ActiveCell.MergeArea.Rows.Count = 10) And (ActiveCell.MergeArea.Columns.Count = 1) Then
                        ' ��������Ă���Z���̃J�E���^�����Z
                        j = j + 1
                        ' �ő�y�[�W�����Z
                        MaxNo = MaxNo + 1
                    ' �Z������������Ă��Ȃ��ꍇ
                    Else
                        ' �ǉ�����y�[�W�����v�Z
                        PageInsertCount = ((myFnameCount - j - BlankCount - 1) \ 6) + 1
                        ' �ǉ�����y�[�W���ő�l�𒴂���ꍇ
                        If (pageNo(ActiveCell.MergeArea.Row, ActiveCell.MergeArea.Column) - 1 + _
                            PageInsertCount * 6) > MaxPageNo Then
                            MsgBox "�ʐ^���ő�y�[�W�𒴂��܂��B" & _
                                vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
                            Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                            Exit Sub
                        End If
                        ' �y�[�W��ǉ�����
                        If AddPages(PageInsertCount) < 0 Then
                            MsgBox "�y�[�W��ǉ��ł��܂���B" & _
                                vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
                            Range(PictureColumn(StartNo) & Format(PictureRow(StartNo))).MergeArea.Select
                            Exit Sub
                        End If
                    End If
                ' �Z������������Ă��鐔����荞�ݖ������Ȃ��Ԃ��肩����
                Loop While j < (myFnameCount - BlankCount)
                ' ���炷�ʐ^�������P���ȏ゠��ꍇ
                If m > 0 Then
                    Application.StatusBar = "�ʐ^�����炵�Ă��܂��B���҂����������B"
                    ' �z��ϐ���錾
                    ReDim PictureName(m - 1)
                    ' �z��̃J�E���^
                    j = 0
                    ' ���ׂĂ̐}�ɑ΂���
                    For Each myPicture In ActiveSheet.Shapes
                        ' �ʐ^�̒ʂ��ԍ������߂�
                        myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
                        ' ���݂̒ʂ��ԍ����傫���ꍇ
                        If myPictureNo >= CurrentNo And _
                            (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                            ' �ʐ^�̖��O��z��ϐ��Ɋi�[
                            PictureName(j) = myPicture.Name
                            ' �z��̃J�E���^�����Z
                            j = j + 1
                        End If
                    Next
                    ' �ʐ^�̖������Q���ȏ�Ȃ�
                    If m > 1 Then
                        ' �z��̏����Ȃ�ׂ���
                        Call QuickSort(PictureName, 0, m - 1)
                    End If
                    ' �ʐ^�����炷
                    For j = m - 1 To 0 Step -1
                        ' �ʐ^�̃y�[�W�ԍ����v�Z
                        myPictureNo = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
                                            ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
                        ' ���炷�y�[�W�ԍ����v�Z
                        NextNo = myPictureNo + myFnameCount - BlankCount
                        ' ���炷�y�[�W�̍s�Ɨ�����߂�
                        NextRow = PictureRow(NextNo)
                        NextColumn = PictureColumn(NextNo)
                        With ActiveSheet.Shapes(PictureName(j))
                            ' �ʐ^�����炷
                            .Top = Range(NextColumn & Format(NextRow)).Top
                            .Left = Range(NextColumn & Format(NextRow)).Left
                            Range(NextColumn & Format(NextRow)).MergeArea.Select
                            ' �ʒu����
                            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
                            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
                        End With
                        ' �B�e���f�[�^�����炷
                        Range(PictureNameBuffer & Format(MinDataBuffer + NextNo)).Value = _
                            Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value
                        Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value = ""
                        Range(PictureDateBuffer & Format(MinDataBuffer + NextNo)).Value = _
                            Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value
                        Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value = ""
                        ' �B�e���\���t���O���n�m�Ȃ�
                        If Range(PictureDateFlag).Value <> 0 Then
                            ' �B�e���e�L�X�g�{�b�N�X�����炷
                            With ActiveSheet.Shapes(PictureDateName(PictureName(j)))
                                .Top = Range(NextColumn & Format(NextRow)).Top
                                .Left = Range(NextColumn & Format(NextRow)).Left
                                Range(NextColumn & Format(NextRow)).MergeArea.Select
                            End With
                            Call PictureDatePosition(PictureName(j))
                        End If
                        ' �R�����g�̂��炷�s�Ɨ���v�Z
                        CurrentRow = PictureRow(myPictureNo)
                        CurrentColumn = CommentColumn(myPictureNo)
                        NextColumn = CommentColumn(NextNo)
                        ' �R�����g�����炷
                        Range(NextColumn & Format(NextRow + 1) & ":" & NextColumn & Format(NextRow + 9)).Value = _
                            Range(CurrentColumn & Format(CurrentRow + 1) & _
                                ":" & CurrentColumn & Format(CurrentRow + 9)).Value
                        ' ���炵���R�����g������
                        Range(CurrentColumn & Format(CurrentRow + 1) & _
                            ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
                    Next j
                    ' �z��ϐ��̉��
                    Erase PictureName
                End If
            End If
        End If
        ' Excel2007�ȍ~�ŁA�摜�������΍�
        myCurrentWindowZoom = ActiveWindow.Zoom
        ' �E�B���h�E�̕\���{�����P�O�O���ɂ���
        ActiveWindow.Zoom = 100
        Application.StatusBar = "�ʐ^����荞��ł��܂��B���҂����������B"
        ' �ʐ^�̎�荞�ݖ��������肩����
       For i = 1 To UBound(SelectedOrder)
    If SelectedOrder(i) > 0 Then
        Set targetCell = PasteStartCell.Offset(SelectedOrder(i) - 1, 0)
        Set myRange = Range(PictureColumn(CurrentNo) & PictureRow(CurrentNo))
        ' Activate�s�v�Ȃ�폜
        Call PastePicture(SelectedPaths(i))
        CurrentNo = CurrentNo + 1
    End If
Next i
        ' �z��ϐ��̉��
        Erase myFname
        ' Excel2007�ȍ~�ŁA�摜�������΍�
        ' �E�B���h�E�̕\���{���𕜌�����
        ActiveWindow.Zoom = myCurrentWindowZoom
    End If
    ' �z��ϐ��̉��
    'Erase myFilename
    ' �ŏ��̃Z����I��
    Set myRange = Range(PictureColumn(StartNo) & Format(PictureRow(StartNo)))
    myRange.MergeArea.Select
    ' �������O�̎ʐ^�J�E���^���u�O�v�ȊO�Ȃ�
    If myFailureCount > 0 Then
        ' �m�F���b�Z�[�W��\��
        If myFailureCount = mySelectedItemsCount Then
            MsgBox "�����ʐ^����荞�ނ��Ƃ͂ł��܂���B", vbOKOnly + vbExclamation, "���m�点"
        Else
            MsgBox Format(mySelectedItemsCount) & "����" & Format(myFailureCount) & "�������ʐ^������܂����B" _
            & vbCrLf & "�����ʐ^�͎�荞�܂�Ă��܂���B�m�F���Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        End If
    End If
    ' �����t�@�C�����̎ʐ^�J�E���^���u�O�v�ȊO�Ȃ�
    If myLongNameCount > 0 Then
        ' �m�F���b�Z�[�W��\��
        If myLongNameCount = mySelectedItemsCount Then
            MsgBox "�ʐ^�̃t�@�C�������������܂�" _
                & vbCrLf & "�t�@�C�����͊g���q�i.JPG�Ȃǁj���܂߂ĂR�O���ȉ��Ƃ��Ă��������B", _
                vbOKOnly + vbExclamation, "���m�点"
        Else
            MsgBox Format(mySelectedItemsCount) & "����" & Format(myLongNameCount) & _
                "�������t�@�C�����̎ʐ^������܂����B" _
                & vbCrLf & "�t�@�C�����͊g���q�i.JPG�Ȃǁj���܂߂ĂR�O���ȉ��Ƃ��Ă��������B" _
                & vbCrLf & "�����t�@�C�����̎ʐ^�͎�荞�܂�Ă��܂���B�m�F���Ă��������B", _
            vbOKOnly + vbExclamation, "���m�点"
        End If
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function AddBlankFunc() As Integer
' �R�}�̒ǉ�����
    Dim i As Integer, j As Integer
    Dim CurrentNo As Integer, PictureNo As Integer, MaxNo As Integer, NextNo As Integer
    Dim CurrentRow As Long, CurrentColumn As String
    Dim NextRow As Long, NextColumn As String
    Dim MaxRow As Long, MaxColumn As String
    Dim myPicture As Shape
    Dim PictureName() As String
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �������I��
        AddBlankFunc = -1
        Exit Function
    End If
    ' �s�A�񂩂�A���݂̃y�[�W�ԍ������߂�
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' ���݂̃y�[�W�ԍ����u�O�v�ȉ��̏ꍇ
    If CurrentNo <= 0 Then
        ' �������I��
        AddBlankFunc = -3
        Exit Function
    End If
    Application.ScreenUpdating = False
    ' �ʐ^�̍ő�y�[�W�ԍ������߂�
    i = 0
    MaxNo = 0
    ' ���ׂĂ̐}�ɑ΂���
    For Each myPicture In ActiveSheet.Shapes
        ' �ŏ��Z���ȍ~�̐}�ɑ΂���
        If myPicture.TopLeftCell.Row >= myMinRow Then
            PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' �ʐ^�̒ʂ��ԍ������݂̎ʐ^���傫���ꍇ
            If PictureNo >= CurrentNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' �ʐ^�̖������J�E���g
                i = i + 1
                ' �ʐ^�̍ő�y�[�W�ԍ����X�V
                If MaxNo < PictureNo Then
                    MaxNo = PictureNo
                End If
            End If
        End If
    Next
    ' �ʐ^�̍ő�y�[�W�ԍ��̎��ɂ���
    MaxNo = MaxNo + 1
    ' �ʐ^���ő�y�[�W�𒴂���ꍇ
    If MaxNo > MaxPageNo Then
        ' �������I��
        AddBlankFunc = -4
        Application.ScreenUpdating = True
        Exit Function
    End If
    ' �ʐ^�̍ő�y�[�W�̍s�Ɨ�����߂�
    MaxRow = PictureRow(MaxNo)
    MaxColumn = PictureColumn(MaxNo)
    ' �y�[�W��ǉ�����
    If (Range(MaxColumn & Format(MaxRow)).MergeArea.Rows.Count <> 10) Or _
        (Range(MaxColumn & Format(MaxRow)).MergeArea.Columns.Count <> 1) Then
        ' �y�[�W��ǉ��ł��Ȃ����
        If AddPages(1) < 0 Then
            ' �������I��
            AddBlankFunc = -2
            Application.ScreenUpdating = True
            Exit Function
        End If
    End If
    ' �ʐ^������ꍇ
    If i > 0 Then
        ' �X�e�[�^�X�o�[�ɏ�Ԃ�\������
        Application.StatusBar = "�ʐ^�����炵�Ă��܂��B���҂����������B"
        ' �z��ϐ���錾
        ReDim PictureName(i - 1)
        j = 0
        ' ���ׂĂ̐}�ɑ΂���
        For Each myPicture In ActiveSheet.Shapes
            ' �ŏ��̍s�ȍ~�̐}�`�ɑ΂���
            If myPicture.TopLeftCell.Row >= myMinRow Then
                ' �ʐ^�̒ʂ��ԍ������߂�
                PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
                ' ���݂̒ʂ��ԍ����傫���ꍇ
                If PictureNo >= CurrentNo And _
                    (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                    ' �ʐ^�̖��O��z��ϐ��Ɋi�[
                    PictureName(j) = myPicture.Name
                    ' �z��̃J�E���^�����Z
                    j = j + 1
                End If
            End If
        Next
        ' �ʐ^�̖������Q���ȏ�Ȃ�
        If i > 1 Then
            ' �z��̏����Ȃ�ׂ���
            Call QuickSort(PictureName, 0, i - 1)
        End If
        ' �ʐ^�����炷
        For j = i - 1 To 0 Step -1
            ' �ʐ^�̍s�Ɨ񂩂�ʐ^�̒ʂ��ԍ������߂�
            PictureNo = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
                                ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
            ' ���炷�ʐ^�̔ԍ��ƍs����ї�����߂�
            NextNo = PictureNo + 1
            NextRow = PictureRow(NextNo)
            NextColumn = PictureColumn(NextNo)
            With ActiveSheet.Shapes(PictureName(j))
                ' �ʐ^�����炷
                .Top = Range(NextColumn & Format(NextRow)).Top
                .Left = Range(NextColumn & Format(NextRow)).Left
                ' ���炵���Z����I��
                Range(NextColumn & Format(NextRow)).MergeArea.Select
                ' �ʐ^�̈ʒu����
                .Top = Selection.Top + ((Selection.Height - .Height) / 2)
                .Left = Selection.Left + ((Selection.Width - .Width) / 2)
            End With
            ' �B�e���\���t���O���n�m�Ȃ�
            If Range(PictureDateFlag).Value <> 0 Then
                ' �B�e���e�L�X�g�{�b�N�X�����炷
                With ActiveSheet.Shapes(PictureDateName(PictureName(j)))
                    .Top = Range(NextColumn & Format(NextRow)).Top
                    .Left = Range(NextColumn & Format(NextRow)).Left
                    Range(NextColumn & Format(NextRow)).MergeArea.Select
                End With
                Call PictureDatePosition(PictureName(j))
            End If
            ' �B�e���f�[�^�����炷
            Range(PictureNameBuffer & Format(MinDataBuffer + NextNo)).Value = _
                Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value
            Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
            Range(PictureDateBuffer & Format(MinDataBuffer + NextNo)).Value = _
                Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value
            Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
            ' �R�����g�̌��݂̍s�Ɨ�����߂�
            CurrentRow = PictureRow(PictureNo)
            CurrentColumn = CommentColumn(PictureNo)
            ' ���炷�R�����g�̗�����߂�
            NextColumn = CommentColumn(NextNo)
            ' �R�����g�����炷
            Range(NextColumn & Format(NextRow + 1) & ":" & NextColumn & Format(NextRow + 9)).Value = _
                Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value
            ' ���炵���Z���̃R�����g������
            Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
        Next j
        ' �z��ϐ��̉��
        Erase PictureName
    End If
    ' ����I��
    AddBlankFunc = 0
    ' �ǉ������]���Z����I��
    Range(PictureColumn(CurrentNo) & Format(PictureRow(CurrentNo))).MergeArea.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Function

Sub AddBlank()
' �R�}�̒ǉ�����
    Select Case AddBlankFunc()
        ' �߂�l�ɂ�鏈��
        Case -1
            MsgBox "�V�[�g���ی삳��Ă��܂��B" _
                & vbCrLf & "�ی���������Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Case -2
            MsgBox "����ȏ�y�[�W��ǉ��ł��܂���B", vbOKOnly + vbExclamation, "���m�点"
        Case -3
            MsgBox "�R�}�ԍ�������������܂���B", vbOKOnly + vbExclamation, "���m�点"
        Case -4
            MsgBox "�ʐ^���ő�y�[�W�𒴂��܂��B" _
                & vbCrLf & "�����𒆒f���܂��B", vbOKOnly + vbExclamation, "���m�点"
    End Select
End Sub

Sub DeleteBlank()
' �R�}�̍폜����
    Dim CurrentNo As Integer, PictureNo As Integer, NextNo As Integer
    Dim CurrentRow As Long, CurrentColumn As String
    Dim NextRow As Long, NextColumn As String
    Dim i As Integer, j As Integer
    Dim myPicture As Shape
    Dim PictureName() As String
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        MsgBox "�V�[�g���ی삳��Ă��܂��B" _
            & vbCrLf & "�ی���������Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        ' �������I��
        Exit Sub
    End If
    ' �s�A�񂩂�A���݂̃y�[�W�ԍ������߂�
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' ���݂̃y�[�W�ԍ����u�O�v�ȉ��̏ꍇ
    If CurrentNo <= 0 Then
        ' �������I��
        Exit Sub
    End If
    Application.ScreenUpdating = False
    ' �ʐ^�̖��������߂�
    i = 0
    For Each myPicture In ActiveSheet.Shapes
        PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        If PictureNo > CurrentNo And _
            (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
            i = i + 1
        End If
    Next
    ' �폜����R�}�̎ʐ^�������[���Ȃ�
    If i = 0 Then
        ' ���݂̃Z���̃R�����g������
        CurrentRow = PictureRow(CurrentNo)
        CurrentColumn = CommentColumn(CurrentNo)
        Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
        ' �B�e���f�[�^������
        Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
        Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Application.StatusBar = "�ʐ^�����炵�Ă��܂��B���҂����������B"
    ' �z��ϐ���錾
    ReDim PictureName(i - 1)
    j = 0
    ' ���ׂĂ̐}�ɑ΂���
    For Each myPicture In ActiveSheet.Shapes
        ' �}�̃y�[�W�ԍ������߂�
        PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' �ʐ^�����݂̃y�[�W�ԍ����傫���ꍇ
        If PictureNo > CurrentNo And _
            (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
            ' �ʐ^�̖��O��z��ϐ��Ɋi�[
            PictureName(j) = myPicture.Name
            ' �z��̃J�E���^�����Z
            j = j + 1
        End If
    Next
    ' �ʐ^�̖������Q���ȏ�Ȃ�
    If i > 1 Then
        ' �z��̏����Ȃ�ׂ���
        Call QuickSort(PictureName, 0, i - 1)
    End If
    ' �ʐ^�����炷
    For j = 0 To i - 1
        PictureNo = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
                            ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
        ' ���炷�ʐ^�̒ʂ��ԍ��ƍs�Ɨ�����߂�
        NextNo = PictureNo - 1
        NextRow = PictureRow(NextNo)
        NextColumn = PictureColumn(NextNo)
        With ActiveSheet.Shapes(PictureName(j))
            ' �ʐ^�����炷
            .Top = Range(NextColumn & Format(NextRow)).Top
            .Left = Range(NextColumn & Format(NextRow)).Left
            Range(NextColumn & Format(NextRow)).MergeArea.Select
            ' �u�`�Q�v�Z���Ŏʐ^���c���łX�O���܂��͂Q�V�O���̏ꍇ�Ɉʒu���߂������΍�
            If NextRow = 2 Then
                Range("A1").RowHeight = TempRowHeight
            End If
            ' �ʐ^�̈ʒu����
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
            ' �u�`�Q�v�Z���̎ʐ^���c���łX�O���܂��͂Q�V�O���̏ꍇ�Ɉʒu���߂������΍�
            If NextRow = 2 Then
                Range("A1").RowHeight = TopRowHeight
            End If
        End With
        ' �B�e���\���t���O���n�m�Ȃ�
        If Range(PictureDateFlag).Value <> 0 Then
            ' �B�e���e�L�X�g�{�b�N�X�����炷
            With ActiveSheet.Shapes(PictureDateName(PictureName(j)))
                .Top = Range(NextColumn & Format(NextRow)).Top
                .Left = Range(NextColumn & Format(NextRow)).Left
                Range(NextColumn & Format(NextRow)).MergeArea.Select
            End With
            Call PictureDatePosition(PictureName(j))
        End If
        ' �B�e���f�[�^�����炷
        Range(PictureNameBuffer & Format(MinDataBuffer + NextNo)).Value = _
            Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value
        Range(PictureNameBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
        Range(PictureDateBuffer & Format(MinDataBuffer + NextNo)).Value = _
            Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value
        Range(PictureDateBuffer & Format(MinDataBuffer + PictureNo)).Value = ""
        ' �R�����g�̌��݂̍s�Ɨ�����߂�
        CurrentRow = PictureRow(PictureNo)
        CurrentColumn = CommentColumn(PictureNo)
        ' ���炷�R�����g�̗�����߂�
        NextColumn = CommentColumn(NextNo)
        ' �R�����g�����炷
        Range(NextColumn & Format(NextRow + 1) & ":" & NextColumn & Format(NextRow + 9)).Value = _
            Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value
        ' ���炵���Z���̃R�����g������
        Range(CurrentColumn & Format(CurrentRow + 1) & ":" & CurrentColumn & Format(CurrentRow + 9)).Value = ""
    Next j
    ' �z��ϐ��̉��
    Erase PictureName
    ' �폜�����Z����I��
    Range(PictureColumn(CurrentNo) & Format(PictureRow(CurrentNo))).MergeArea.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function pageNo(myRow As Long, myColumn As Long) As Integer
' �s�ԍ��A��ԍ�����A�ʐ^�̒ʂ��ԍ������߂�
    If myColumn = 1 Then
        pageNo = (myRow - myMinRow) \ 11 + 1
    Else
        pageNo = 0
    End If
End Function

Function PictureRow(PictureNo As Integer) As Long
' �ʐ^�̒ʂ��ԍ�����A�s�ԍ������߂�
    PictureRow = CLng(PictureNo - 1) * 11 + myMinRow
End Function

Function PictureColumn(PictureNo As Integer) As String
' �ʐ^�̒ʂ��ԍ�����A��ԍ������߂�
    PictureColumn = "A"
End Function

Function CommentColumn(PictureNo As Integer) As String
' �ʐ^�̒ʂ��ԍ�����A�R�����g���̗�ԍ������߂�
    CommentColumn = "B"
End Function

Sub QuickSort(ByRef PictureName() As String, ByVal ArrayMinNo As Integer, ByVal ArrayMaxNo As Integer)
' �z��̕��בւ������i�N�C�b�N�\�[�g�j
    Dim MinNo As Integer
    Dim MaxNo As Integer
    Dim MidValue As Integer
    Dim TempValue As String
    ' �����̒l���擾
    MidValue = pageNo(ActiveSheet.Shapes(PictureName(Int((ArrayMinNo + ArrayMaxNo) / 2))).TopLeftCell.Row, _
                ActiveSheet.Shapes(PictureName(Int((ArrayMinNo + ArrayMaxNo) / 2))).TopLeftCell.Column)
    MinNo = ArrayMinNo
    MaxNo = ArrayMaxNo
    ' ���肩����
    Do
        ' �z��̍ŏ����̂��肩����
        Do
            ' �z��̒l�������̒l���傫���ꍇ
            If (pageNo(ActiveSheet.Shapes(PictureName(MinNo)).TopLeftCell.Row, _
                    ActiveSheet.Shapes(PictureName(MinNo)).TopLeftCell.Column) >= MidValue) Then
                ' ���肩�����𔲂���
                Exit Do
            End If
            ' �ŏ������P���Z
            MinNo = MinNo + 1
        Loop
        ' �z��̍ő呤�̂��肩����
        Do
            ' �z��̒l�������̒l��菬�����ꍇ
            If (MidValue >= pageNo(ActiveSheet.Shapes(PictureName(MaxNo)).TopLeftCell.Row, _
                            ActiveSheet.Shapes(PictureName(MaxNo)).TopLeftCell.Column)) Then
                ' ���肩�����𔲂���
                Exit Do
            End If
            ' �ő呤���P���Z
            MaxNo = MaxNo - 1
        Loop
        ' �ŏ������ő呤���傫���Ȃ�Ώ����I��
        If (MinNo >= MaxNo) Then
            Exit Do
        End If
        ' �z��̓���ւ�
        TempValue = PictureName(MinNo)
        PictureName(MinNo) = PictureName(MaxNo)
        PictureName(MaxNo) = TempValue
        ' �ŏ������P���Z
        MinNo = MinNo + 1
        ' �ő呤���P���Z
        MaxNo = MaxNo - 1
    Loop
    ' �z��̍ŏ������ċA�ŕ��בւ�
    If (ArrayMinNo < MinNo - 1) Then
        Call QuickSort(PictureName, ArrayMinNo, MinNo - 1)
    End If
    ' �z��̍ő呤���ċA�ŕ��בւ�
    If (MaxNo + 1 < ArrayMaxNo) Then
        Call QuickSort(PictureName, MaxNo + 1, ArrayMaxNo)
    End If
End Sub

Sub PictureNumbering()
' �ʐ^���̔ԍ��t�ԏ���
    Dim myPage As Integer, i As Integer, j As Integer, k As Integer
    Dim myPicture As Shape, PictureName() As String, PictureNo As Integer
    Dim myRange As Range, CurrentRange As Range, CurrentRow As Long, CurrentColumn As Long
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        MsgBox "�V�[�g���ی삳��Ă��܂��B" _
            & vbCrLf & "�ی���������Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        ' �������I��
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.StatusBar = "�ʐ^���ɔԍ���U�蒼���Ă��܂��B���҂����������B"
    ' ���݂̃Z�����L��
    Set CurrentRange = Cells(ActiveCell.Row, ActiveCell.Column)
    ' �ʐ^�̖����J�E���^
    i = 0
    ' �ʐ^�̖��������߂�
    For Each myPicture In ActiveSheet.Shapes
        PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' �}���ʐ^�Ńy�[�W�ԍ����u�O�v���傫���ꍇ
        If PictureNo > 0 And (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
            ' �������J�E���g
            i = i + 1
        End If
    Next
    ' �ʐ^���P���ȏ゠��ꍇ
    If i > 0 Then
        ' �z��ϐ���錾
        ReDim PictureName(i - 1)
        ' �z��ϐ��p�J�E���^
        j = 0
        ' ���ׂĂ̐}�ɑ΂���
        For Each myPicture In ActiveSheet.Shapes
            PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' �}���ʐ^�Ńy�[�W�ԍ����u�O�v���傫���ꍇ
            If PictureNo > 0 And (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' �z��ɐ}�̖��O����
                PictureName(j) = myPicture.Name
                ' �J�E���^�����Z
                j = j + 1
            End If
        Next
        ' �ʐ^���Q���ȏ�Ȃ�
        If i > 1 Then
            ' �z��̕��בւ�
            Call QuickSort(PictureName, 0, i - 1)
        End If
    End If
    ' �ʐ^�̒ʂ��ԍ�
    j = 0
    ' �R�}�̒ʂ��ԍ�
    k = 1
    ' �y�[�W�ԍ�
    myPage = 1
    ' �擪�̃R�}�̃Z����I��
    Set myRange = Range(PictureColumn(1) & Format(PictureRow(1)))
    myRange.Select
    ' ���肩����
    Do
        ' �ʐ^���P���ȏ�̏ꍇ
        If i > 0 And j < i Then
            ' �ʐ^�̍s�Ɨ�����߂�
            CurrentRow = ActiveSheet.Shapes(PictureName(j)).TopLeftCell.MergeArea.Row
            CurrentColumn = ActiveSheet.Shapes(PictureName(j)).TopLeftCell.MergeArea.Column
            ' ���݂̃Z���̍s�Ɨ�Ǝʐ^�̍s�Ɨ񂪈�v�����ꍇ
            If myRange.Row = CurrentRow And myRange.Column = CurrentColumn Then
                ' �y�[�W�ԍ����Z���ɏ�������
                Range(CommentColumn(k) & Format(PictureRow(k))).Value = "No." & Format(myPage)
                 ' �y�[�W�����Z
                myPage = myPage + 1
               ' ���̎ʐ^��
                j = j + 1
            ' ��v���Ȃ��ꍇ
            Else
                ' �y�[�W�ԍ�������
                Range(CommentColumn(k) & Format(PictureRow(k))).Value = ""
            End If
        ' �ŏI�y�[�W�܂�
        Else
            ' �y�[�W�ԍ�������
            Range(CommentColumn(k) & Format(PictureRow(k))).Value = ""
        End If
        ' ���̃R�}�ֈړ�
        k = k + 1
        Set myRange = Range(PictureColumn(k) & Format(PictureRow(k)))
        myRange.MergeArea.Select
    ' �Z�����������Ă���Ԃ��肩����
    Loop While (myRange.MergeArea.Rows.Count = 10) And (k <= MaxPageNo)
    ' �z��ϐ��̉��
    Erase PictureName
    ' ���݂̃Z����I��
    CurrentRange.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub SerialNumbering()
' �R�}���Ƃ̔ԍ��t�ԏ���
    Dim myPage As Integer, k As Integer
    Dim myRange As Range, CurrentRange As Range
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        MsgBox "�V�[�g���ی삳��Ă��܂��B" _
            & vbCrLf & "�ی���������Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        ' �������I��
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.StatusBar = "�R�}���ɔԍ���U�蒼���Ă��܂��B���҂����������B"
    ' ���݂̃Z�����L��
    Set CurrentRange = Cells(ActiveCell.Row, ActiveCell.Column)
    ' �R�}�̒ʂ��ԍ�
    k = 1
    ' �y�[�W�ԍ�
    myPage = 1
    ' �擪�̃R�}�̃Z����I��
    Set myRange = Range(PictureColumn(1) & Format(PictureRow(1)))
    myRange.Select
    ' ���肩����
    Do
        ' �y�[�W�ԍ����Z���ɏ�������
        Range(CommentColumn(k) & Format(PictureRow(k))).Value = "No." & Format(myPage)
        ' �y�[�W�����Z
        myPage = myPage + 1
        ' ���̃R�}�ֈړ�
        k = k + 1
        Set myRange = Range(PictureColumn(k) & Format(PictureRow(k)))
        myRange.MergeArea.Select
    ' �Z�����������Ă���Ԃ��肩����
    Loop While (myRange.MergeArea.Rows.Count = 10) And (k <= MaxPageNo)
    ' ���݂̃Z����I��
    CurrentRange.Select
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Function PictureExist(CurrentRow As Long, CurrentColumn As Long) As Boolean
' �Z���Ɏʐ^���\��t�����Ă��邩�m�F����֐�
    Dim CurrentNo As Integer, PictureNo As Integer
    Dim myPicture As Shape
    ' ���݂̃Z���̃y�[�W�ԍ������߂�
    CurrentNo = pageNo(CurrentRow, CurrentColumn)
    ' ���ׂĂ̎ʐ^�ɂ��ČJ��Ԃ�
    For Each myPicture In ActiveSheet.Shapes
        If myPicture.TopLeftCell.Row >= myMinRow Then
            ' �ʐ^�̃y�[�W�ԍ������߂�
            PictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' �ʐ^�̃y�[�W�ԍ��ƌ��݂̃Z���̃y�[�W�ԍ������������
            If PictureNo = CurrentNo And _
                (myPicture.Type = msoPicture Or myPicture.Type = msoLinkedPicture) Then
                ' �߂�l���uTrue�v�ɂ���
                PictureExist = True
                ' �������I��
                Exit Function
            End If
        End If
    Next
    ' ���݂̃Z���Ɏʐ^���\��t�����Ă��Ȃ���Ζ߂�l���uFalse�v�ɂ���
    PictureExist = False
End Function

Sub MoveToEnd()
' �ʐ^�𖖔��Ɉړ������鏈��
    Dim myPicture As Shape
    Dim myRange As Range
    Dim myPictureNo As Integer, myMaxNo As Integer
    Dim CommentData(10) As String, i As Integer, AdjustNo As Integer
    Dim myPictureName As String, myPictureDate As String, myDate As String, myType As Integer
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �}�`�̍ő�y�[�W��
    myMaxNo = 0
    ' �������O�̃J�E���^
    i = 0
    ' �ړ���ƈړ����̃V�[�g���قȂ��
    If SwapSourceSheet <> SwapDestSheet Then
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
    End If
    ' ���ׂĂ̐}�`�ɑ΂���
    For Each myPicture In ActiveSheet.Shapes
        ' �}�`�̍���Z�����ŏ��Z���ȏ�̏ꍇ
        If myPicture.TopLeftCell.Row >= myMinRow Then
            ' �}�`�̃y�[�W�������߂�
            myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
            ' �y�[�W�̍ő吔�����߂�
            If myMaxNo < myPictureNo Then
                myMaxNo = myPictureNo
            End If
        End If
        ' �ړ���ƈړ����̃V�[�g���قȂ��
        If SwapSourceSheet <> SwapDestSheet Then
            ' �ʐ^�̖��O���ړ����̎ʐ^�̖��O�Ɠ����Ȃ�
            If myPicture.Name = SwapSourceName Then
                ' �J�E���^�����Z
                i = i + 1
                ' �J�E���^���P�ȏ�Ȃ�
                If i > 0 Then
                    ' ���b�Z�[�W��\�����ď������I��
                    MsgBox "�ړ���̃V�[�g�ɓ����ʐ^������܂��B" & vbCrLf & "�����ʐ^��\��t���邱�Ƃ͂ł��܂���B", _
                        vbOKOnly + vbExclamation, "���m�点"
                    Exit Sub
                End If
            End If
        End If
    Next
    ' �ړ���̃V�[�g��I��
    Worksheets(SwapDestSheet).Select
    ' �ړ����ƈړ���̃V�[�g���قȂ�ꍇ
    If SwapSourceSheet <> SwapDestSheet Then
        ' �y�[�W�ԍ��̕␳�l���{�P
        AdjustNo = 1
        ' �ő�y�[�W�̃Z������������Ă��Ȃ����
        If Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo))).MergeArea.Rows.Count <> 10 Or _
            Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo))).MergeArea.Columns.Count <> 1 Then
            ' �y�[�W��ǉ�����
            If AddPages(1) < 0 Then
                ' �y�[�W�̒ǉ����ł��Ȃ��ꍇ���b�Z�[�W��\�����ď������I��
                MsgBox "�y�[�W��ǉ��ł��܂���ł����B" & vbCrLf & "�������I�����܂��B", vbOKOnly + vbExclamation, "���m�点"
                Exit Sub
            End If
        End If
    ' �ړ����ƈړ���̃V�[�g�������ꍇ
    Else
        ' �y�[�W�ԍ��̕␳�l���O
        AdjustNo = 0
    End If
    ' �ړ����V�[�g��I��
    Worksheets(SwapSourceSheet).Select
    ' �R�����g���̃f�[�^��ۑ�
    For i = 1 To 9
        CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
    Next i
    ' �B�e���f�[�^��ۑ�
    myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag).Value <> 0 Then
        ' �B�e���e�L�X�g�{�b�N�X���폜
        ActiveSheet.Shapes(PictureDateName(SwapSourceName)).Delete
    End If
    ' �ʐ^��؂���
    ActiveSheet.Shapes(SwapSourceName).Cut
    ' �R�}�̍폜
    Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
    Call DeleteBlank
    ' �ړ���̃V�[�g��I��
    Worksheets(SwapDestSheet).Select
    ' �ړ����ƈړ���̃V�[�g���قȂ�ꍇ
    If SwapSourceSheet <> SwapDestSheet Then
        ' �y�[�W�ԍ��̕␳�l���{�P
        AdjustNo = 1
    ' �ړ����ƈړ���̃V�[�g�������ꍇ
    Else
        ' �y�[�W�ԍ��̕␳�l���O
        AdjustNo = 0
    End If
    ' �ŏI�y�[�W�Ɏʐ^��\��t��
    Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo))).MergeArea.Select
    ActiveSheet.Paste
    With ActiveSheet.Shapes(SwapSourceName)
        ' �ʐ^��\��t�����Z����I��
        .TopLeftCell.MergeArea.Select
        ' �ʐ^�̏c������Œ�
        .LockAspectRatio = msoTrue
        ' �ʐ^�̊p�x���X�O���܂��͂Q�V�O���̏ꍇ
        If .Rotation = 90 Or .Rotation = 270 Then
        ' �ʐ^�̕����Z���̍����ɂ��낦��
            .Width = Int(ActiveCell.MergeArea.Height) - 4
            ' �ʐ^�̍������Z���̕��ɂ��낦��
            If .Height > Int(ActiveCell.MergeArea.Width) - 4 Then
                .Height = Int(ActiveCell.MergeArea.Width) - 4
            End If
        ' �ʐ^�̊p�x���O���܂��͂P�W�O���̏ꍇ
        Else
            ' �ʐ^�̍������Z���̍����ɂ��낦��
            .Height = Int(ActiveCell.MergeArea.Height) - 4
            ' �ʐ^�̕����Z���̕��ɂ��낦��
            If .Width > Int(ActiveCell.MergeArea.Width) - 4 Then
                .Width = Int(ActiveCell.MergeArea.Width) - 4
            End If
        End If
        ' �ʐ^�̈ʒu����
        .Top = Selection.Top + ((Selection.Height - .Height) / 2)
        .Left = Selection.Left + ((Selection.Width - .Width) / 2)
    End With
    ' �R�����g���̃f�[�^���ő�y�[�W�Ɉڂ�
    For i = 1 To 9
        Range(CommentColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo) + i)).Value = CommentData(i)
    Next i
    ' �B�e���f�[�^���ő�y�[�W�Ɉڂ�
    Range(PictureNameBuffer & Format(MinDataBuffer + myMaxNo + AdjustNo)).Value = myPictureName
    Range(PictureDateBuffer & Format(MinDataBuffer + myMaxNo + AdjustNo)).Value = myPictureDate
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag).Value <> 0 Then
        myDate = myPictureDate
        ' �B�e���̕\���`���t���O���G���R�[�h
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        ' �B�e���̃e�L�X�g�{�b�N�X����}
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If
    ' �ő�y�[�W�̃Z����I��
    Set myRange = Range(PictureColumn(myMaxNo + AdjustNo) & Format(PictureRow(myMaxNo + AdjustNo)))
    myRange.MergeArea.Select
End Sub

Sub MoveToHere()
' �ʐ^���ړ������鏈��
    Dim SwapSourceRow As Long, SwapSourceColumn As String
    Dim SwapDestRow As Long, SwapDestColumn As String
    Dim CommentData(10) As String, i As Integer
    Dim AdjustNo As Integer
    Dim myPicture As Shape
    Dim myPictureName As String, myPictureDate As String, myDate As String, myType As Integer
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �ʐ^�̈ړ����̍s�Ɨ���擾
    SwapSourceRow = PictureRow(SwapSourceNo)
    SwapSourceColumn = PictureColumn(SwapSourceNo)
    ' �ʐ^�̈ړ���̍s�Ɨ���擾
    SwapDestRow = PictureRow(SwapDestNo)
    SwapDestColumn = PictureColumn(SwapDestNo)
    ' �ړ���̃V�[�g�ƈړ����̃V�[�g���قȂ�ꍇ
    If SwapSourceSheet <> SwapDestSheet Then
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
        ' �����ʐ^�̃J�E���^
        i = 0
        ' ���ׂĂ̐}�ɑ΂��ČJ��Ԃ�
        For Each myPicture In ActiveSheet.Shapes
            ' �ʐ^�̖��O���ړ����̎ʐ^�̖��O�Ɠ����Ȃ�
            If myPicture.Name = SwapSourceName Then
                ' �J�E���^�����Z
                i = i + 1
                ' �J�E���^���P�ȏ�Ȃ�
                If i > 0 Then
                    ' ���b�Z�[�W��\�����ď������I��
                    MsgBox "�ړ���̃V�[�g�ɓ����ʐ^������܂��B" & vbCrLf & "�����ʐ^��\��t���邱�Ƃ͂ł��܂���B", _
                        vbOKOnly + vbExclamation, "���m�点"
                    Exit Sub
                End If
            End If
        Next
        ' �ړ���̃y�[�W��I��
        Range(SwapDestColumn & Format(SwapDestRow)).MergeArea.Select
        ' �]���R�}��ǉ�
        If AddBlankFunc < 0 Then
            ' ���b�Z�[�W��\�����ď������I��
            MsgBox "�]���R�}��ǉ��ł��܂���ł����B" & vbCrLf & "�������I�����܂��B", vbOKOnly + vbExclamation, "���m�点"
            Exit Sub
        End If
    End If
    ' �ړ����̃V�[�g��I��
    Worksheets(SwapSourceSheet).Select
    ' �ړ����̃Z����I��
    ActiveSheet.Shapes(SwapSourceName).TopLeftCell.MergeArea.Select
    ' �ړ����̃R�����g���̃f�[�^��ۑ�
    For i = 1 To 9
        CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
    Next i
    ' �ړ����̎B�e���f�[�^��ۑ�
    myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
    Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = ""
    Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = ""
    ' �ړ����̃Z����I��
    ActiveSheet.Shapes(SwapSourceName).TopLeftCell.MergeArea.Select
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag).Value <> 0 Then
        ' �B�e���e�L�X�g�{�b�N�X���폜
        ActiveSheet.Shapes(PictureDateName(SwapSourceName)).Delete
    End If
    ' �ړ����̃Z���̎ʐ^��؂���
    ActiveSheet.Shapes(SwapSourceName).Cut
    ' �ړ����̗]���R�}���폜
    Call DeleteBlank
    ' �ړ���̃y�[�W�ԍ����ړ����y�[�W�ԍ����傫���ꍇ
    If (SwapSourceNo < SwapDestNo) And (SwapSourceSheet = SwapDestSheet) Then
        ' �ړ���y�[�W�ԍ��̕␳�l���|�P�ɂ���
        AdjustNo = -1
    ' �ړ���̃y�[�W�ԍ����ړ����y�[�W�ԍ���菬�����ꍇ
    Else
        ' �ړ���y�[�W�ԍ��̕␳�l���O�ɂ���
        AdjustNo = 0
    End If
    ' �ړ���ƈړ����̃V�[�g�������Ȃ�
    If SwapSourceSheet = SwapDestSheet Then
        ' �ړ���̃Z����I��
        Range(PictureColumn(SwapDestNo + AdjustNo) & Format(PictureRow(SwapDestNo + AdjustNo))).MergeArea.Select
        ' �]���R�}��ǉ�
        If AddBlankFunc < 0 Then
            ' ���b�Z�[�W��\�����ď������I��
            MsgBox "�]���R�}��ǉ��ł��܂���ł����B" & vbCrLf & "�������I�����܂��B", vbOKOnly + vbExclamation, "���m�点"
            Exit Sub
        End If
    ' �ړ���ƈړ����̃V�[�g���قȂ�ꍇ
    Else
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
    End If
    ' �ʐ^�̈ړ���Ɉړ����̎ʐ^��\��t��
    Range(PictureColumn(SwapDestNo + AdjustNo) & Format(PictureRow(SwapDestNo + AdjustNo))).MergeArea.Select
    ActiveSheet.Paste
    With ActiveSheet.Shapes(SwapSourceName)
        ' �ʐ^��\��t�����Z����I��
        .TopLeftCell.MergeArea.Select
        ' �ʐ^�̏c������Œ�
        .LockAspectRatio = msoTrue
        ' �ʐ^�̊p�x���X�O���܂��͂Q�V�O���̏ꍇ
        If .Rotation = 90 Or .Rotation = 270 Then
        ' �ʐ^�̕����Z���̍����ɂ��낦��
            .Width = Int(ActiveCell.MergeArea.Height) - 4
            ' �ʐ^�̍������Z���̕��ɂ��낦��
            If .Height > Int(ActiveCell.MergeArea.Width) - 4 Then
                .Height = Int(ActiveCell.MergeArea.Width) - 4
            End If
        ' �ʐ^�̊p�x���O���܂��͂P�W�O���̏ꍇ
        Else
            ' �ʐ^�̍������Z���̍����ɂ��낦��
            .Height = Int(ActiveCell.MergeArea.Height) - 4
            ' �ʐ^�̕����Z���̕��ɂ��낦��
            If .Width > Int(ActiveCell.MergeArea.Width) - 4 Then
                .Width = Int(ActiveCell.MergeArea.Width) - 4
            End If
        End If
        ' �ʐ^�̈ʒu����
        .Top = Selection.Top + ((Selection.Height - .Height) / 2)
        .Left = Selection.Left + ((Selection.Width - .Width) / 2)
    End With
    ' �R�����g���̃f�[�^���ړ���ֈڂ�
    For i = 1 To 9
        Range(CommentColumn(SwapDestNo + AdjustNo) & Format(PictureRow(SwapDestNo + AdjustNo) + i)).Value = CommentData(i)
    Next i
    ' �B�e���f�[�^���ړ���ֈڂ�
    Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo + AdjustNo)).Value = myPictureName
    Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo + AdjustNo)).Value = myPictureDate
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag).Value <> 0 Then
        myDate = myPictureDate
        ' �B�e���̕\���`���t���O���G���R�[�h
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        ' �B�e���̃e�L�X�g�{�b�N�X����}
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If
    ' �z��ϐ��̉��
    Erase CommentData
End Sub

Sub ExchangePicture()
' �ʐ^�����ւ��鏈��
    Dim SwapSourceRow As Long, SwapSourceColumn As String
    Dim SwapDestRow As Long, SwapDestColumn As String
    Dim CommentData(10) As String, i As Integer
    Dim myPicture As Shape
    Dim myPictureName As String, myPictureDate As String, myDestPictureName As String, myDestPictureDate As String
    Dim SwapSourceDate As String, SwapDestDate As String, myDate As String, myType As Integer
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �ʐ^�̓���ւ����̍s�Ɨ���擾
    SwapSourceRow = PictureRow(SwapSourceNo)
    SwapSourceColumn = PictureColumn(SwapSourceNo)
    ' �ʐ^�̓���ւ���̍s�Ɨ���擾
    SwapDestRow = PictureRow(SwapDestNo)
    SwapDestColumn = PictureColumn(SwapDestNo)
    ' �ړ���ƈړ����̃V�[�g�������ꍇ
    If SwapSourceSheet = SwapDestSheet Then
        ' �ʐ^�̈ړ���Ɉړ����̎ʐ^���ړ�
        With ActiveSheet.Shapes(SwapSourceName)
            .Top = Range(SwapDestColumn & Format(SwapDestRow)).Top
            .Left = Range(SwapDestColumn & Format(SwapDestRow)).Left
            Range(SwapDestColumn & Format(SwapDestRow)).Select
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        End With
        ' �ʐ^�̈ړ����Ɉړ���̎ʐ^���ړ�
        With ActiveSheet.Shapes(SwapDestName)
            .Top = Range(SwapSourceColumn & Format(SwapSourceRow)).Top
            .Left = Range(SwapSourceColumn & Format(SwapSourceRow)).Left
            Range(SwapSourceColumn & Format(SwapSourceRow)).Select
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        End With
        ' �ړ����̃R�����g���̃f�[�^���擾
        For i = 1 To 9
            CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
        Next i
        ' �R�����g���̃f�[�^���ړ��悩��ړ����Ɉڂ�
        Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 1) & ":" & _
            CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 9)).Value = _
        Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 1) & ":" & _
            CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 9)).Value
        ' �R�����g���̃f�[�^���ړ���ֈڂ�
        For i = 1 To 9
            Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + i)).Value = CommentData(i)
        Next i
        ' �ړ����̎B�e���f�[�^���擾
        myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        ' �B�e���f�[�^���ړ��悩��ړ����ֈڂ�
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = _
            Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = _
            Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        ' �B�e���f�[�^���ړ���ֈڂ�
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureName
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureDate
        SwapSourceDate = PictureDateName(SwapSourceName)
        SwapDestDate = PictureDateName(SwapDestName)
        ' �B�e���\���t���O���n�m�Ȃ�
        If Range(PictureDateFlag).Value <> 0 Then
            ' �ʐ^�̈ړ���Ɉړ����̎B�e�����ړ�
            With ActiveSheet.Shapes(SwapSourceDate)
                .Top = Range(SwapDestColumn & Format(SwapDestRow)).Top
                .Left = Range(SwapDestColumn & Format(SwapDestRow)).Left
                Range(SwapDestColumn & Format(SwapDestRow)).Select
            End With
            Call PictureDatePosition(SwapSourceName)
            ' �ʐ^�̈ړ����Ɉړ���̎B�e�����ړ�
            With ActiveSheet.Shapes(SwapDestDate)
                .Top = Range(SwapSourceColumn & Format(SwapSourceRow)).Top
                .Left = Range(SwapSourceColumn & Format(SwapSourceRow)).Left
                Range(SwapSourceColumn & Format(SwapSourceRow)).Select
            End With
            Call PictureDatePosition(SwapDestName)
        End If
    ' �ړ���ƈړ����̃V�[�g���قȂ�ꍇ
    Else
        ' �ړ����̃V�[�g��I��
        Worksheets(SwapSourceSheet).Select
        ' �����ʐ^�̃J�E���^
        i = 0
        ' ���ׂĂ̐}�ɑ΂��ČJ��Ԃ�
        For Each myPicture In ActiveSheet.Shapes
            ' �ʐ^�̖��O���ړ���̎ʐ^�̖��O�Ɠ����Ȃ�
            If myPicture.Name = SwapDestName Then
                ' �J�E���^�����Z
                i = i + 1
                ' �J�E���^���P�ȏ�Ȃ�
                If i > 0 Then
                    ' ���b�Z�[�W��\�����ď������I��
                    MsgBox "�ړ���̃V�[�g�ɓ����ʐ^������܂��B" & vbCrLf & "�����ʐ^��\��t���邱�Ƃ͂ł��܂���B", _
                        vbOKOnly + vbExclamation, "���m�点"
                    Exit Sub
                End If
            End If
        Next
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
        ' �����ʐ^�̃J�E���^
        i = 0
        ' ���ׂĂ̐}�ɑ΂��ČJ��Ԃ�
        For Each myPicture In ActiveSheet.Shapes
            ' �ʐ^�̖��O���ړ���̎ʐ^�̖��O�Ɠ����Ȃ�
            If myPicture.Name = SwapSourceName Then
                ' �J�E���^�����Z
                i = i + 1
                ' �J�E���^���P�ȏ�Ȃ�
                If i > 0 Then
                    ' ���b�Z�[�W��\�����ď������I��
                    MsgBox "�ړ���̃V�[�g�ɓ����ʐ^������܂��B" & vbCrLf & "�����ʐ^��\��t���邱�Ƃ͂ł��܂���B", _
                        vbOKOnly + vbExclamation, "���m�点"
                    Exit Sub
                End If
            End If
        Next
        ' �ړ����̃V�[�g��I��
        Worksheets(SwapSourceSheet).Select
        ' �ړ����A�ړ���̎B�e���f�[�^�̖��O��ݒ�
        SwapSourceDate = PictureDateName(SwapSourceName)
        SwapDestDate = PictureDateName(SwapDestName)
        ' �B�e���\���t���O���n�m�Ȃ�
        If Range(PictureDateFlag).Value <> 0 Then
            ' �ړ����̎B�e���f�[�^���폜
            ActiveSheet.Shapes(SwapSourceDate).Delete
        End If
        ' �ړ����̎ʐ^��؂���
        ActiveSheet.Shapes(SwapSourceName).Cut
        ' �ړ����̃R�����g���̃f�[�^���擾
        For i = 1 To 9
            CommentData(i) = Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + i)).Value
        Next i
        ' �ړ����̎B�e���f�[�^��ۑ�
        myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        myPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
        ' �ʐ^�̈ړ���Ɉړ����̎ʐ^��\��t��
        Range(PictureColumn(SwapDestNo) & Format(PictureRow(SwapDestNo))).MergeArea.Select
        ActiveSheet.Paste
        With ActiveSheet.Shapes(SwapSourceName)
            ' �ʐ^��\��t�����Z����I��
            .TopLeftCell.MergeArea.Select
            ' �ʐ^�̏c������Œ�
            .LockAspectRatio = msoTrue
            ' �ʐ^�̊p�x���X�O���܂��͂Q�V�O���̏ꍇ
            If .Rotation = 90 Or .Rotation = 270 Then
            ' �ʐ^�̕����Z���̍����ɂ��낦��
                .Width = Int(ActiveCell.MergeArea.Height) - 4
                ' �ʐ^�̍������Z���̕��ɂ��낦��
                If .Height > Int(ActiveCell.MergeArea.Width) - 4 Then
                    .Height = Int(ActiveCell.MergeArea.Width) - 4
                End If
            ' �ʐ^�̊p�x���O���܂��͂P�W�O���̏ꍇ
            Else
                ' �ʐ^�̍������Z���̍����ɂ��낦��
                .Height = Int(ActiveCell.MergeArea.Height) - 4
                ' �ʐ^�̕����Z���̕��ɂ��낦��
                If .Width > Int(ActiveCell.MergeArea.Width) - 4 Then
                    .Width = Int(ActiveCell.MergeArea.Width) - 4
                End If
            End If
            ' �ʐ^�̈ʒu����
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        End With
        ' �B�e���\���t���O���n�m�Ȃ�
        If Range(PictureDateFlag).Value <> 0 Then
            myDate = myPictureDate
            ' �B�e���̕\���`���t���O���G���R�[�h
            myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
            ' �B�e���̃e�L�X�g�{�b�N�X����}
            Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
        End If
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
        ' �B�e���\���t���O���n�m�Ȃ�
        If Range(PictureDateFlag).Value <> 0 Then
            ' �ړ���̎B�e���f�[�^���폜
            ActiveSheet.Shapes(SwapDestDate).Delete
        End If
        ' �ړ���̎ʐ^��؂���
        ActiveSheet.Shapes(SwapDestName).Cut
        ' �R�����g���̃f�[�^���ړ��悩��ړ����ֈڂ�
        Worksheets(SwapSourceSheet).Range(CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 1) & ":" & _
            CommentColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo) + 9)).Value = _
            Worksheets(SwapDestSheet).Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 1) & ":" & _
            CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + 9)).Value
        ' �ړ���̎B�e���f�[�^��ۑ�
        myDestPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        myDestPictureDate = Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value
        ' �B�e���f�[�^���ړ���ֈڂ�
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureName
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapDestNo)).Value = myPictureDate
        ' �ړ����̃V�[�g��I��
        Worksheets(SwapSourceSheet).Select
        ' �ʐ^�̈ړ����Ɉړ���̎ʐ^��\��t��
        Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
        ActiveSheet.Paste
        With ActiveSheet.Shapes(SwapDestName)
            ' �ʐ^��\��t�����Z����I��
            .TopLeftCell.MergeArea.Select
            ' �ʐ^�̏c������Œ�
            .LockAspectRatio = msoTrue
            ' �ʐ^�̊p�x���X�O���܂��͂Q�V�O���̏ꍇ
            If .Rotation = 90 Or .Rotation = 270 Then
            ' �ʐ^�̕����Z���̍����ɂ��낦��
                .Width = Int(ActiveCell.MergeArea.Height) - 4
                ' �ʐ^�̍������Z���̕��ɂ��낦��
                If .Height > Int(ActiveCell.MergeArea.Width) - 4 Then
                    .Height = Int(ActiveCell.MergeArea.Width) - 4
                End If
            ' �ʐ^�̊p�x���O���܂��͂P�W�O���̏ꍇ
            Else
                ' �ʐ^�̍������Z���̍����ɂ��낦��
                .Height = Int(ActiveCell.MergeArea.Height) - 4
                ' �ʐ^�̕����Z���̕��ɂ��낦��
                If .Width > Int(ActiveCell.MergeArea.Width) - 4 Then
                    .Width = Int(ActiveCell.MergeArea.Width) - 4
                End If
            End If
            ' �ʐ^�̈ʒu����
            .Top = Selection.Top + ((Selection.Height - .Height) / 2)
            .Left = Selection.Left + ((Selection.Width - .Width) / 2)
        End With
        ' �B�e���\���t���O���n�m�Ȃ�
        If Range(PictureDateFlag).Value <> 0 Then
            myDate = myDestPictureDate
            ' �B�e���̕\���`���t���O���G���R�[�h
            myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
            ' �B�e���̃e�L�X�g�{�b�N�X����}
            Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
        End If
        ' �B�e���f�[�^���ړ����ֈڂ�
        Range(PictureNameBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = myDestPictureName
        Range(PictureDateBuffer & Format(MinDataBuffer + SwapSourceNo)).Value = myDestPictureDate
        ' �ړ���̃V�[�g��I��
        Worksheets(SwapDestSheet).Select
        ' �R�����g���̃f�[�^���ړ���Ɉڂ�
        For i = 1 To 9
            Range(CommentColumn(SwapDestNo) & Format(PictureRow(SwapDestNo) + i)).Value = CommentData(i)
        Next i
    End If
    ' �z��ϐ��̉��
    Erase CommentData
    ' �ړ���̃V�[�g��I��
    Worksheets(SwapDestSheet).Select
End Sub

Sub SaveWOMacro()
' �}�N���Ȃ��ŕۑ��I������
    Dim myWorksheet As Worksheet
    Dim myShape As Shape
    Dim myFname As String, Ans As Integer, mySheet As String
    ' ���݂̃V�[�g�����擾
    mySheet = ActiveSheet.Name
    ' ���݂̃u�b�N���ύX����Ă���ꍇ
    If ActiveWorkbook.Saved = False Then
        ' �m�F���b�Z�[�W��\��
        Ans = MsgBox("���݂̃u�b�N�͕ύX����Ă��܂��B" & vbCrLf & "�ύX��ۑ����Ă����܂����H", _
            vbYesNoCancel + vbExclamation, "�m�F")
        ' �L�����Z�����N���b�N���ꂽ�ꍇ
        If Ans = vbCancel Then
            ' �������I��
            Exit Sub
        ' �x�d�r���N���b�N���ꂽ�ꍇ
        ElseIf Ans = vbYes Then
            ' ���݂̃u�b�N��ۑ�
            ActiveWorkbook.Save
        End If
    End If
    ' �u�b�N�̃t�@�C���l�[�����擾���Ċg���q����菜��
    myFname = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    ' �J��Ԃ�
    Do
        ' �t�@�C�����̓���
        myFname = InputBox("�t�@�C��������͂��Ă��������B", "�}�N�����폜���ăt�@�C����ۑ����I�����܂��B", myFname)
        ' �L�����Z�����N���b�N���ꂽ�ꍇ
        If myFname = "" Then
            ' �������I��
            Exit Sub
        End If
        ' �t�@�C�����̃`�F�b�N
        If InStr(myFname, "*") > 0 Or InStr(myFname, "?") > 0 Or InStr(myFname, "\") > 0 Or _
            InStr(myFname, ":") > 0 Or InStr(myFname, "<") > 0 Or InStr(myFname, ">") > 0 Or _
            InStr(myFname, "[") > 0 Or InStr(myFname, "]") > 0 Or InStr(myFname, "|") > 0 Or _
            InStr(myFname, Chr(34)) > 0 Then
            ' ���b�Z�[�W��\��
            Ans = MsgBox("�t�@�C�������s���ł��B�ȉ��̕����͎g���܂���B" & vbCrLf & _
                " * ? : < > [ ] | \ " & Chr(34), vbOKOnly + vbCritical, "����")
        ' �t�@�C�����̕������`�F�b�N
        ElseIf Len(ActiveWorkbook.Path & "\" & myFname & ".xlsx") > 218 Then
            ' ���b�Z�[�W��\��
            Ans = MsgBox("�t�@�C�������������܂��B", vbOKOnly + vbCritical, "����")
        ' �����t�@�C�����̃u�b�N������ꍇ
        ElseIf Dir(ActiveWorkbook.Path & "\" & myFname & ".xlsx") <> "" Then
            ' ���b�Z�[�W��\��
            Ans = MsgBox(myFname & ".xlsx�͊��ɑ��݂��܂��B" & vbCrLf & "�㏑�����܂����H", _
                vbYesNoCancel + vbExclamation, "�m�F")
        Else
            Ans = vbYes
        End If
        ' �L�����Z�����N���b�N���ꂽ�ꍇ
        If Ans = vbCancel Then
            ' �������I��
            Exit Sub
        End If
    ' �x�d�r���N���b�N����܂ŌJ��Ԃ�
    Loop While Ans <> vbYes
    ' �u�b�N���̂��ׂẴ��[�N�V�[�g�ɂ��ČJ��Ԃ�
    For Each myWorksheet In ThisWorkbook.Worksheets
        ' ���[�N�V�[�g��I��
        myWorksheet.Select
        ' ���[�N�V�[�g���̂��ׂĂ̐}�`�ɂ��ČJ��Ԃ�
        For Each myShape In ActiveSheet.Shapes
            ' �}�`�̃}�N�����폜
            If myShape.OnAction <> "" Then
                myShape.OnAction = ""
            End If
        Next
    Next
    ' �V�[�g��I��
    Worksheets(mySheet).Select
    ' �}�N���Ȃ��Ńt�@�C����ۑ�
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & myFname, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    ' �G�N�Z�����I��
    Application.Quit
End Sub

Function PictureDate(myPic As String) As String
' �ʐ^�̎B�e�������擾����
    Dim objFS, objFile, shellObj, folderObj, MyFile, myPath, i As Long, GetDetailsNumber As Long

    Set objFS = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFS.GetFile(myPic)
    MyFile = objFile.Name
    myPath = Replace(myPic, MyFile, "")
    myPath = Left(myPath, Len(myPath) - 1)
    Set shellObj = CreateObject("Shell.Application")
    Set folderObj = shellObj.Namespace(myPath)

    GetDetailsNumber = -1
    ' Exif���̍��ڔԍ����擾
    For i = 0 To 100 '�\���ɑ傫������
        If folderObj.GetDetailsOf("", i) = "�B�e����" Then
            GetDetailsNumber = i
            Exit For
        End If
    Next i
    ' �B�e�������擾���A�߂�l�ɂ���
    PictureDate = folderObj.GetDetailsOf(folderObj.ParseName(MyFile), GetDetailsNumber)
    
    Set objFS = Nothing
    Set objFile = Nothing
    Set shellObj = Nothing
    Set folderObj = Nothing
End Function

Sub PictureDateOFF()
' �B�e������������
    Dim i As Integer, j As Integer, k As Integer, myShape As Shape, PictureName() As String
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �ʐ^�̖��������߂�
    i = 0
    For Each myShape In ActiveSheet.Shapes
        If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
            i = i + 1
        End If
    Next
    ' �ʐ^���P���ȏ�Ȃ�
    If i > 0 Then
        ' ���I�z��̐錾
        ReDim PictureName(i - 1)
        j = 0
        ' ���ׂĂ̐}�ɑ΂���
        For Each myShape In ActiveSheet.Shapes
            ' �}���ʐ^�Ȃ�
            If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
                ' �ʐ^�̖��O���擾
                PictureName(j) = myShape.Name
                ' �ʐ^�̖������J�E���g
                j = j + 1
            End If
        Next
        ' �S�Ă̐}�ɑ΂���
        For k = 0 To j - 1
            ' �B�e���e�L�X�g�{�b�N�X���폜
            ActiveSheet.Shapes(PictureDateName(PictureName(k))).Delete
        Next k
    End If
    ' �B�e���\���t���O�����Z�b�g
    Range(PictureDateFlag).Value = ""
    Erase PictureName
End Sub

Sub PictureDateON()
' �B�e�����\������
    Dim i As Integer, j As Integer, k As Integer, myShape As Shape, PictureName() As String, myDate As String
    Dim CurrentRow As Long, CurrentColumn As Long, myType As Integer
    CurrentRow = ActiveCell.Row
    CurrentColumn = ActiveCell.Column
    ' �ʐ^�̖��������߂�
    i = 0
    For Each myShape In ActiveSheet.Shapes
        If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
            i = i + 1
        End If
    Next
    ' �ʐ^���P���ȏ�Ȃ�
    If i > 0 Then
        ' ���I�z��̐錾
        ReDim PictureName(i - 1)
        j = 0
        ' ���ׂĂ̐}�ɑ΂���
        For Each myShape In ActiveSheet.Shapes
            ' �}���ʐ^�Ȃ�
            If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
                ' �ʐ^�̖��O���擾
                PictureName(j) = myShape.Name
                ' �ʐ^�̖������J�E���g
                j = j + 1
            End If
        Next
        ' �S�Ă̐}�ɑ΂���
        For k = 0 To j - 1
            ' �B�e���f�[�^���擾
            myDate = Range(PictureDateBuffer & Format(MinDataBuffer + pageNo(ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Row, _
                ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Column))).Value
            ' �}�̃Z����I��
            Cells(ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Row, ActiveSheet.Shapes(PictureName(k)).TopLeftCell.Column).Select
            ' �B�e���̕\���`���t���O���G���R�[�h
            myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
            ' �B�e���̃e�L�X�g�{�b�N�X����}
            Call AddPictureDate(PictureName(k), PictureDateFormat(myDate, myType))
        Next k
    End If
    ' �B�e���\���t���O���Z�b�g
    Range(PictureDateFlag).Value = 1
    Erase PictureName
    ' ���݂̃Z����I��
    Cells(CurrentRow, CurrentColumn).Select
End Sub

Sub AddPictureDate(ByVal myPictureName As String, ByVal myDate As String)
' �B�e���̃e�L�X�g�{�b�N�X����}
    Dim WeekLen As Integer
    ' �j���\������ꍇ�̕������␳
    If Range(WeekDisp).Value <> 0 And Range(WeekLang).Value = 0 Then
        WeekLen = 2
    Else
        WeekLen = 1
    End If
    ' �e�L�X�g�{�b�N�X����}
    With ActiveSheet.Shapes.AddTextbox _
        (msoTextOrientationHorizontal, _
        ActiveCell.Left, ActiveCell.Top, ((Len(myDate) + WeekLen) * Range(DateFontSize).Value / 2), (Range(DateFontSize).Value + DateHeightOffset))
        ' �e�L�X�g�{�b�N�X�̖��O��ݒ�
        .Name = PictureDateName(myPictureName)
        ' �e�L�X�g����
        .TextFrame2.TextRange.Characters.Text = myDate
        ' �t�H���g�T�C�Y
        .TextFrame2.TextRange.Characters.Font.Size = Range(DateFontSize).Value
        ' �t�H���g�̐F
        .TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = _
            RGB(Range(DateFontColorR).Value, Range(DateFontColorG).Value, Range(DateFontColorB).Value)
        ' �����ɂ���
        .TextFrame2.TextRange.Characters.Font.Bold = (Range(DateFontBold).Value <> 0)
        ' �����̍��}�[�W��
        .TextFrame2.MarginLeft = 0
        ' �����̉E�}�[�W��
        .TextFrame2.MarginRight = 0
        ' �����̏�}�[�W��
        .TextFrame2.MarginTop = 0
        ' �����̉��}�[�W��
        .TextFrame2.MarginBottom = 0
        ' �����̐܂�Ԃ��Ȃ�
        .TextFrame2.WordWrap = msoFalse
        ' �e�L�X�g�{�b�N�X�̕������E�l��
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        ' �e�L�X�g�{�b�N�X�̕����̏㉺�𒆉����낦
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        ' ������
        .Line.Visible = False
        ' �h��Ԃ�����
        .Fill.Visible = False
        ' �ʒu����
        .Top = Selection.Top
        .Left = Selection.Left
        ' �e�L�X�g�{�b�N�X�ɁA�}�N����o�^
        .OnAction = "PictureDateClick"
    End With
    ' �e�L�X�g�{�b�N�X�̈ʒu����
    Call PictureDatePosition(myPictureName)
End Sub

Sub PictureDatePosition(ByVal myPictureName As String)
' �B�e���̃e�L�X�g�{�b�N�X�̈ʒu����
    Dim PictureWidth As Double, PictureHeight As Double
    '�ʐ^�̕��A�������擾
    PictureWidth = ActiveSheet.Shapes(myPictureName).Width
    PictureHeight = ActiveSheet.Shapes(myPictureName).Height
    ' �ʐ^�̃Z����I��
    ActiveSheet.Shapes(myPictureName).TopLeftCell.MergeArea.Select
    ' �ʐ^���X�O���܂��͂Q�V�O����]���Ă���ꍇ
    If ((ActiveSheet.Shapes(myPictureName).Rotation = 90 Or ActiveSheet.Shapes(myPictureName).Rotation = 270)) Then
        ' �B�e�����ʐ^�̉E���Ɉʒu����
        With ActiveSheet.Shapes(PictureDateName(myPictureName))
            .Top = Selection.Top + (Selection.Height - PictureWidth) / 2 + PictureWidth - .Height - Range(DateYOffset).Value
            .Left = Selection.Left + (Selection.Width - PictureHeight) / 2 + PictureHeight - .Width - Range(DateXOffset).Value
        End With
    ' �ʐ^���O���܂��͂P�W�O����]���Ă���ꍇ
    Else
        ' �B�e�����ʐ^�̉E���Ɉʒu����
        With ActiveSheet.Shapes(PictureDateName(myPictureName))
            .Top = Selection.Top + (Selection.Height - PictureHeight) / 2 + PictureHeight - .Height - Range(DateYOffset).Value
            .Left = Selection.Left + (Selection.Width - PictureWidth) / 2 + PictureWidth - .Width - Range(DateXOffset).Value
        End With
    End If
End Sub

Function PictureDateFormat(ByVal myDate As String, ByVal myType As Integer) As String
' �B�e���̃t�H�[�}�b�g�ݒ�
    If myDate <> "" Then
        ' ���p�����ɕϊ�
        myDate = StrConv(myDate, vbNarrow)
        ' �\���ł��Ȃ��������폜
        myDate = Replace(myDate, "?", "")
        ' �j���\���u�i�v������ꍇ�A�N�����Ǝ����𒊏o
        If (InStr(myDate, "(") > 8) And (InStr(myDate, "(") <= 11) Then
            myDate = Left(myDate, InStr(myDate, "(") - 1) & " " & Right(myDate, Len(myDate) - InStrRev(myDate, " "))
        End If
    End If
    ' �B�e���\���`�������t�����̏ꍇ�imyType�̒l���f�R�[�h���Ĕ�r�j
    If (myType Mod 2) <> 0 Then
        ' �����������f�[�^�Ȃ�
        If IsDate(myDate) Then
            ' �\���`������t�݂̂ɂ���
            myDate = Format(myDate, "yyyy/mm/dd")
            ' �j����\������ꍇ�imyType�̒l���f�R�[�h���Ĕ�r�j
            If ((myType \ 4) Mod 2) <> 0 Then
                ' �j�����p��ŕ\������ꍇ�imyType�̒l���f�R�[�h���Ĕ�r�j
                If ((myType \ 8) Mod 2) <> 0 Then
                    myDate = Format(myDate, "yyyy/mm/dd(ddd)")
                ' �j������{��ŕ\��
                Else
                    myDate = Format(myDate, "yyyy/mm/dd(aaa)")
                End If
            End If
        ' ����myDate�������e�[�^�Ƃ��Ĉ����Ȃ��ꍇ
        Else
            ' ���������ŕ\��
            myDate = "****/**/**"
        End If
    ' �B�e���\���`�������t�Ǝ����̏ꍇ
    Else
        ' �����������f�[�^�Ȃ�
        If IsDate(myDate) Then
            ' �\���`������t�Ǝ����ɂ���
            myDate = Format(myDate, "yyyy/mm/dd h:nn")
            ' �j����\������ꍇ�imyType�̒l���f�R�[�h���Ĕ�r�j
            If ((myType \ 4) Mod 2) <> 0 Then
                ' �j�����p��ŕ\������ꍇ�imyType�̒l���f�R�[�h���Ĕ�r�j
                If ((myType \ 8) Mod 2) <> 0 Then
                    myDate = Format(myDate, "yyyy/mm/dd(ddd) h:nn")
                ' �j������{��ŕ\��
                Else
                    myDate = Format(myDate, "yyyy/mm/dd(aaa) h:nn")
                End If
            End If
        ' ����myDate�������f�[�^�Ƃ��Ĉ����Ȃ��ꍇ
        Else
            ' ���t�Ǝ��������ŕ\��
            myDate = "****/**/** **:**"
        End If
    End If
    ' �B�e���̓��t��؂�L���Ɂu�D�v���g���ꍇ�imyType�̒l���f�R�[�h���Ĕ�r�j
    If ((myType \ 2) Mod 2) <> 0 Then
        ' ���t�́u/�v���u.�v�ɒu��������
        myDate = Replace(myDate, "/", ".")
    End If
    ' �ϊ��������t������߂�l�ɂ���
    PictureDateFormat = myDate
End Function

Function PictureDateName(ByVal myPictureName As String) As String
' �B�e���e�L�X�g�{�b�N�X�̖��O�𐶐�
    ' �ʐ^�̖��O�̍Ō�̂P�������uDateSuffix�v�Œ�`���Ă��镶���ɒu��������
    PictureDateName = Left(myPictureName, (Len(myPictureName) - 1)) & DateSuffix
End Function

Sub PictureDateClick()
' �B�e���e�L�X�g�{�b�N�X���N���b�N�����ꍇ
    ' ���݂̃V�[�g��I��
    ActiveSheet.Select
    ' ���[�U�[�t�H�[���X���Ăяo��
    UserForm9.Show vbModal
End Sub

Function GetPictureNoFromDate(ByVal myName As String) As Integer
' �B�e���e�L�X�g�{�b�N�X�̖��O����ʐ^�̃y�[�W�ԍ������߂�
    Dim i As Integer, j As Integer, myShape As Shape, PictureName() As String, myPictureDate As String
    ' �ʐ^�̖��������߂�
    i = 0
    For Each myShape In ActiveSheet.Shapes
        If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
            i = i + 1
        End If
    Next
    ' �ʐ^���P���ȏ�Ȃ�
    If i > 0 Then
        ' ���I�z��̐錾
        ReDim PictureName(i - 1)
        j = 0
        ' ���ׂĂ̐}�ɑ΂���
        For Each myShape In ActiveSheet.Shapes
            ' �}���ʐ^�Ȃ�
            If myShape.Type = msoPicture Or myShape.Type = msoLinkedPicture Then
                ' �ʐ^�̖��O���擾
                PictureName(j) = myShape.Name
                ' �ʐ^�̖������J�E���g
                j = j + 1
            End If
        Next
        ' ���ׂĂ̎ʐ^�̖��O�ɑ΂���
        For j = 0 To i - 1
            ' �B�e���e�L�X�g�{�b�N�X�̖��O�������ƈ�v�����
            If PictureDateName(PictureName(j)) = myName Then
                ' �J��Ԃ��𔲂���
                Exit For
            End If
        Next j
        ' �ʐ^�̃y�[�W�ԍ���߂�l�ɐݒ�
        GetPictureNoFromDate = pageNo(ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Row, _
            ActiveSheet.Shapes(PictureName(j)).TopLeftCell.Column)
    ' �ʐ^�������ꍇ��
    Else
        ' �߂�l���[���ɂ��Ă���
        GetPictureNoFromDate = 0
    End If
    ' �z��ϐ��̉��
    Erase PictureName
End Function

Sub PictureDateDispSequence()
    Dim Ans As Integer
    ' �B�e�����\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag).Value <> 0 Then
        Ans = MsgBox("�B�e�������������܂����H" & vbCrLf & _
            "��������ꍇ�́u�͂��v��" & vbCrLf & _
            "�B�e�����̕\����ύX����ꍇ�́u�������v���N���b�N���Ă��������B", vbYesNoCancel + vbInformation, "�m�F")
        If Ans = vbYes Then
            ' �B�e���������������Ăяo��
            Call PictureDateOFF
        ElseIf Ans = vbNo Then
            ' �B�e�����̕\���I��
            UserForm8.Show vbModal
        End If
    ' �B�e�����\���t���O���n�e�e�Ȃ�
    Else
        ' �B�e�����̕\���I��
        UserForm8.Show vbModal
    End If
End Sub

Sub DeleteLastPages()
' �����̃y�[�W�폜����
    Dim myBottomCount As Long
    Dim myPictureNo As Integer, myMaxNo As Integer
    Dim myPicture As Shape
    Application.ScreenUpdating = False
    Application.StatusBar = "�����̃y�[�W���폜���Ă��܂��B���҂����������B"
    ' �ʐ^�̍ő�ʂ��ԍ�
    myMaxNo = 1
    ' ���ׂĂ̐}�ɂ�������
    For Each myPicture In ActiveSheet.Shapes
        myPictureNo = pageNo(myPicture.TopLeftCell.Row, myPicture.TopLeftCell.Column)
        ' �}�̍s�����y�[�W�̍ő�l���傫���ꍇ
        If myPicture.BottomRightCell.Row > MaxPageRow Then
            ' ���b�Z�[�W��\������
            MsgBox "�ʐ^���ő�y�[�W�𒴂��ē\��t�����Ă��܂��B", vbOKOnly + vbExclamation, "���m�点"
            ' �������I��
            Application.ScreenUpdating = True
            Application.StatusBar = False
            Exit Sub
        ElseIf (myPictureNo > myMaxNo) Then
            ' �}�̍ő�ʂ��ԍ������߂�
            myMaxNo = myPictureNo
        End If
    Next
    ' �ŏI�s���v�Z
    myBottomCount = (PictureRow(myMaxNo) \ 33) * 33 + 33
    ' �y�[�W���폜
    If myBottomCount < 65536 Then
        Range("A" & Format(myBottomCount + 1) & ":G65536").EntireRow.Delete
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ' �ʐ^�̍ő�y�[�W��I��
    Range(PictureColumn(myMaxNo) & Format(PictureRow(myMaxNo))).Select
End Sub
