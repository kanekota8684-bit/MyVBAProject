VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "������I�����Ă��������B"
   ClientHeight    =   2540
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' ����]�{�^��
    ' ��]�����i�}�̖��O�A��]�p�x�j���Ăяo��
    Call RotationProc(Application.Caller, -90)
    ' �B�e���e�L�X�g�{�b�N�X�̈ʒu����
    If Range(PictureDateFlag).Value <> 0 Then
        Call PictureDatePosition(Application.Caller)
    End If
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �E��]�{�^��
    ' ��]�����i�}�̖��O�A��]�p�x�j���Ăяo��
    Call RotationProc(Application.Caller, 90)
    ' �B�e���e�L�X�g�{�b�N�X�̈ʒu����
    If Range(PictureDateFlag).Value <> 0 Then
        Call PictureDatePosition(Application.Caller)
    End If
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' �L�����Z���{�^��
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton4_Click()
' �؂���{�^��
    Dim CurrentNo As Integer, CurrentName As String
    On Error Resume Next
    CurrentName = ActiveSheet.Shapes(Application.Caller).Name
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag) <> 0 Then
        ' �B�e���e�L�X�g�{�b�N�X���폜
        ActiveSheet.Shapes(PictureDateName(CurrentName)).Delete
    End If
    ' �}��؂���
    ActiveSheet.Shapes(Application.Caller).Cut
    ' �B�e���f�[�^��؂���o�b�t�@�֕ۑ�
    Range(PictureNameBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value
    Range(PictureDateBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value
    ' �B�e���f�[�^������
    Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    ' ���݂̃V�[�g�����L��
    CutDataSheet = ActiveSheet.Name
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton5_Click()
' �����ɗ]���R�}��ǉ��{�^��
    Me.Hide
    ' �R�}�̒ǉ����Ăяo��
    Call AddBlank
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton6_Click()
' �ʐ^���폜���ċl�߂�{�^��
    Dim CurrentNo As Integer, CurrentName As String
    On Error Resume Next
    CurrentName = ActiveSheet.Shapes(Application.Caller).Name
    CurrentNo = pageNo(ActiveCell.Row, ActiveCell.Column)
    Me.Hide
    ' �B�e���\���t���O���n�m�Ȃ�
    If Range(PictureDateFlag) <> 0 Then
        ' �B�e���e�L�X�g�{�b�N�X���폜
        ActiveSheet.Shapes(PictureDateName(CurrentName)).Delete
    End If
    ' �}��؂���
    ActiveSheet.Shapes(Application.Caller).Cut
    ' �B�e���f�[�^��؂���o�b�t�@�Ɉڂ�
    Range(PictureNameBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value
    Range(PictureDateBuffer & Format(CutDataBuffer)).Value = _
        Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value
    ' �B�e���f�[�^������
    Range(PictureNameBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    Range(PictureDateBuffer & Format(MinDataBuffer + CurrentNo)).Value = ""
    ' ���݂̃V�[�g���L��
    CutDataSheet = ActiveSheet.Name
    ' �R�}�̍폜���Ăяo��
    Call DeleteBlank
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton7_Click()
' �����Ɏ捞�{�^��
    Me.Hide
    ' �����Ɉꊇ�捞���Ăяo��
    Call GetMultiPictureFromHere
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton8_Click()
' �y�[�W�ǉ��{�^��
    Me.Hide
    ' �y�[�W�ǉ����Ăяo���i�����͒ǉ������j
    Call AddPageProc(1)
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton9_Click()
' ����{�^��
    Me.Hide
    ' ����v���r���[���Ăяo��
    UserForm1.Show vbModal
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton10_Click()
' �R�}���ɔԍ��U�����{�^��
    Me.Hide
    ' �R�}���̔ԍ��t�ԏ������Ăяo��
    Call SerialNumbering
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton11_Click()
' �ʐ^���ɔԍ��U�����{�^��
    Me.Hide
    ' �ʐ^���̔ԍ��t�ԏ������Ăяo��
    Call PictureNumbering
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton12_Click()
' �����Ɏ捞�{�^��
    Me.Hide
    ' �ꊇ�捞���Ăяo��
    Call GetMultiPicture
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton13_Click()
' ���̎ʐ^�����ւ��{�^��
    ' �ʐ^�̈ړ����̃y�[�W�ԍ��Ǝʐ^�̖��O�ƌ��݂̃V�[�g����ݒ�
    SwapSourceNo = pageNo(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row, ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column)
    SwapSourceName = ActiveSheet.Shapes(Application.Caller).Name
    SwapSourceSheet = ActiveSheet.Name
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' ���[�U�[�t�H�[���T�����[�h���X�ŕ\��
    UserForm5.Show vbModeless
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton14_Click()
' �B�e���\���{�^��
    Me.Hide
    ' �B�e�����\���V�[�P���X���Ăяo��
    Call PictureDateDispSequence
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏����ݒ�
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �{�^���𖳌��ɂ���
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
        ' �������I��
        Exit Sub
    End If
    With ActiveSheet.Shapes(Application.Caller)
        ' �ʐ^�̓\��t�����Ă���Z�����P�O��������Ă��Ďʐ^�̊p�x���X�O�����݂̏ꍇ
        If (.TopLeftCell.MergeArea.Count = 10) And (.TopLeftCell.MergeArea.Columns.Count = 1) And _
            ((.Rotation = 0) Or (.Rotation = 90) Or (.Rotation = 180) Or (.Rotation = 270)) Then
            ' �ʐ^�̃Z����I������
            .TopLeftCell.MergeArea.Select
        ' �ʐ^�̓\��t�����Ă���Z�����P�O��������Ă��Ȃ��܂��͎ʐ^�̊p�x���X�O�����݂łȂ��ꍇ
        Else
            ' �{�^���𖳌��ɂ���
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
    ' �B�e���\���{�^������ݒ�
    If Range(PictureDateFlag).Value <> 0 Then
        CommandButton14.Caption = "�B�e���\���n�e�e(F)"
    Else
        CommandButton14.Caption = "�B�e���\���n�m(F)"
    End If
End Sub
