VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "������I�����Ă��������B"
   ClientHeight    =   1130
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4128
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' �����Ɏʐ^���ړ��{�^��
    Call MoveToHere
    ' �ʐ^�̈ړ����ƈړ���y�[�W�ԍ����N���A
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' ���̎ʐ^�Ɠ���ւ��{�^��
    Call ExchangePicture
    ' �ʐ^�̈ړ����ƈړ���y�[�W�ԍ����N���A
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' �L�����Z���{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' �ʐ^�̈ړ����y�[�W�̃Z����I��
    Worksheets(SwapSourceSheet).Select
    Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
    ' ���[�U�[�t�H�[���T�����[�h���X�ŕ\��
    UserForm5.Show vbModeless
    ' ���[�U�[�t�H�[�������
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏�������
    ' ���N���b�N���ꂽ�ʐ^�̃Z����I��
    ActiveSheet.Shapes(Application.Caller).TopLeftCell.MergeArea.Select
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �{�^���𖳌��ɂ���
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ���[�U�[�t�H�[���̏I���O�̏���
    ' �u�~�v�ŕ��悤�Ƃ����ꍇ
    If CloseMode = vbFormControlMenu Then
        ' ���[�U�[�t�H�[�����\���ɂ���
        Me.Hide
        ' �ʐ^�̈ړ����y�[�W�̃Z����I��
        Worksheets(SwapSourceSheet).Select
        Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
        ' ���[�U�[�t�H�[���T�����[�h���X�ŕ\��
        UserForm5.Show vbModeless
    End If
End Sub
