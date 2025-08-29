VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "������I�����Ă��������B"
   ClientHeight    =   3020
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   4800
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' �����Ɏ捞�{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' �����Ɉꊇ�捞���Ăяo��
    Call GetMultiPictureFromHere
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �y�[�W�ǉ��{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' �y�[�W�ǉ����Ăяo���i�����͒ǉ������j
    Call AddPageProc(1)
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' ����{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' ����v���r���[���Ăяo��
    UserForm1.Show vbModal
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton4_Click()
' �L�����Z���{�^��
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton5_Click()
' �\��t���{�^��
    Dim myPictureName As String
    Dim myPicture As Shape
    Dim i As Integer, CurrentNo As Integer, myDate As String, myType As Integer
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' ��ʕ\���̍X�V�����Ȃ��悤�ɂ���
    Application.ScreenUpdating = False
    ' �\��t��
    ActiveSheet.Paste
    ' �ʐ^���X���Ă�����X�O�����݂ɕ␳����
    With Selection.ShapeRange
        If .Rotation < 45 Or .Rotation >= 315 Then
            .Rotation = 0
        ElseIf .Rotation >= 45 And .Rotation < 135 Then
            .Rotation = 90
        ElseIf .Rotation >= 135 And .Rotation < 225 Then
            .Rotation = 180
        ElseIf .Rotation >= 225 And .Rotation < 315 Then
            .Rotation = 270
        End If
        ' �ʐ^���Ŕw�ʂɈړ�
        .ZOrder msoSendToBack
        ' �ʐ^�̖��O���擾
        myPictureName = .Name
    End With
    ' �������O�̃J�E���^
    i = 0
    ' ���ׂĂ̐}�ɑ΂���
    For Each myPicture In ActiveSheet.Shapes
        ' �������O�������
        If myPicture.Name = myPictureName Then
            ' �J�E���^�����Z
            i = i + 1
            ' �������O�̎ʐ^�������
            If i > 1 Then
                ' �ʐ^��؂���
                Selection.Cut
                ' ���[�U�[�t�H�[�����\���ɂ���
                Me.Hide
                ' ���b�Z�[�W��\�����ē\��t���𒆒f������
                MsgBox "�����ʐ^��\��t���邱�Ƃ͂ł��܂���B", vbOKOnly + vbExclamation, "���m�点"
                ' ���[�U�[�t�H�[�����������
                Application.ScreenUpdating = True
                Unload Me
                Exit Sub
            End If
        End If
    Next
    With ActiveSheet.Shapes(myPictureName)
        ' �ʐ^��\��t�����Z����I��
        .TopLeftCell.MergeArea.Select
        ' �u�`�Q�v�Z���Ŏʐ^���c���łX�O���܂��͂Q�V�O���̏ꍇ�Ɉʒu���߂������΍�
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TempRowHeight
        End If
        ' �ʐ^�̏c������Œ肷��
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
        ' �u�`�Q�v�Z���Ŏʐ^���c���łX�O���܂��͂Q�V�O���̏ꍇ�Ɉʒu���߂������΍�
        If Selection.Row = 2 Then
            Range("A1").RowHeight = TopRowHeight
        End If
    End With
    ' �B�e���f�[�^�𕜌�
    Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = _
        Worksheets(CutDataSheet).Range(PictureNameBuffer & Format(CutDataBuffer)).Value
    Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = _
        Worksheets(CutDataSheet).Range(PictureDateBuffer & Format(CutDataBuffer)).Value
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag).Value <> 0 Then
        ' �B�e�����擾
        myDate = Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value
        ' �B�e���̕\���`���t���O���G���R�[�h
        myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
        ' �B�e���̃e�L�X�g�{�b�N�X����}
        Call AddPictureDate(myPictureName, PictureDateFormat(myDate, myType))
    End If
    ' ��ʕ\���̍X�V��������
    Application.ScreenUpdating = True
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton6_Click()
' �����Ɏ捞�{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' �ꊇ��荞�݂��Ăяo��
    Call GetMultiPicture
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton7_Click()
' �]���R�}���߂�{�^��
    Me.Hide
    ' �R�}�̍폜���Ăяo��
    Call DeleteBlank
    Unload Me
End Sub

Private Sub CommandButton8_Click()
' �]���R�}��ǉ��{�^��
    Me.Hide
    ' �R�}�̒ǉ����Ăяo��
    Call AddBlank
    Unload Me
End Sub

Private Sub CommandButton9_Click()
' �R�}���ɔԍ��U�����{�^��
    Me.Hide
    ' �R�}���̔ԍ��t�ԏ������Ăяo��
    Call SerialNumbering
    Unload Me
End Sub

Private Sub CommandButton10_Click()
' �ʐ^���ɔԍ��U�����{�^��
    Me.Hide
    ' �ʐ^���̔ԍ��t�ԏ������Ăяo��
    Call PictureNumbering
    Unload Me
End Sub

Private Sub CommandButton11_Click()
' �V���[�g�J�b�g�\���{�^��
    ' �V���[�g�J�b�g�\���t���O���Z�b�g
    ShortCutFlag = False
    Unload Me
End Sub

Private Sub CommandButton12_Click()
' �}�N���Ȃ��ŕۑ��I���{�^��
    Me.Hide
    ' �}�N���Ȃ��ŕۑ��I���������Ăяo��
    Call SaveWOMacro
    Unload Me
End Sub

Private Sub CommandButton13_Click()
' �B�e���\���{�^��
    Me.Hide
    ' �B�e�����\���V�[�P���X���Ăяo��
    Call PictureDateDispSequence
    Unload Me
End Sub

Private Sub CommandButton14_Click()
' �����̃y�[�W�폜�{�^��
    Me.Hide
    ' �����̃y�[�W�폜�������Ăяo��
    Call DeleteLastPages
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏����ݒ�
    Dim PictureExistStatus As Boolean
    ' �V���[�g�J�b�g�\���t���O�����Z�b�g
    ShortCutFlag = True
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �{�^���𖳌��ɂ���
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton9.Enabled = False
        CommandButton10.Enabled = False
        CommandButton12.Enabled = False
        CommandButton13.Enabled = False
        CommandButton14.Enabled = False
        ' �������I��
        Exit Sub
    End If
    ' �u�`��v�̃Z������������Ă���ꍇ
    If (ActiveCell.MergeArea.Rows.Count = 10) And (ActiveCell.MergeArea.Columns.Count = 1) And _
        (ActiveCell.Row <= 65463) And (ActiveCell.Column = 1) Then
        ' ���݂̃y�[�W�ԍ����u�O�v�ȉ��̏ꍇ
        If pageNo(ActiveCell.Row, ActiveCell.Column) <= 0 Then
            ' �����Ɏ捞�{�^���𖳌��ɂ���
            CommandButton1.Enabled = False
            ' �\�t�{�^���𖳌��ɂ���
            CommandButton5.Enabled = False
            ' �]���R�}���߂�{�^���𖳌��ɂ���
            CommandButton7.Enabled = False
            ' �]���R�}��ǉ��{�^���𖳌��ɂ���
            CommandButton8.Enabled = False
        ' ���݂̃y�[�W�ԍ����u�P�v�ȏ�̏ꍇ
        Else
            ' ���݂̃Z���Ɏʐ^������ꍇ
            PictureExistStatus = PictureExist(ActiveCell.Row, ActiveCell.Column)
            ' �]���R�}���߂�{�^���𖳌��ɂ���
            CommandButton7.Enabled = Not PictureExistStatus
            ' �N���b�v�{�[�h���}�`�̏ꍇ
            If Application.ClipboardFormats(1) = xlClipboardFormatPICT Then
                ' ���݂̃Z���Ɏʐ^������ꍇ�͓\�t�{�^���𖳌��ɂ���
                CommandButton5.Enabled = Not PictureExistStatus
            ' �N���b�v�{�[�h���}�`�łȂ��ꍇ
            Else
                ' �\�t�{�^���𖳌��ɂ���
                CommandButton5.Enabled = False
            End If
        End If
    ' �Z������������Ă��Ȃ��ꍇ
    Else
        ' �����Ɏ捞�{�^���𖳌��ɂ���
        CommandButton1.Enabled = False
        ' �\�t�{�^���𖳌��ɂ���
        CommandButton5.Enabled = False
        ' �]���R�}���߂�{�^���𖳌��ɂ���
        CommandButton7.Enabled = False
        ' �]���R�}��ǉ��{�^���𖳌��ɂ���
        CommandButton8.Enabled = False
    End If
    ' �ʐ^��؂��肵�����̃V�[�g�����ݒ肳��Ă��Ȃ����
    If CutDataSheet = "" Then
        ' �\��t���{�^���𖳌��ɂ���
        CommandButton5.Enabled = False
    End If
    ' �B�e���\���{�^������ݒ�
    If Range(PictureDateFlag).Value <> 0 Then
        CommandButton13.Caption = "�B�e���\���n�e�e(F)"
    Else
        CommandButton13.Caption = "�B�e���\���n�m(F)"
    End If
End Sub
