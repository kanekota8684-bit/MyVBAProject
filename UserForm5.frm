VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "�ʐ^�̓���ւ��������ł��B���̑�������Ȃ��ł��������B"
   ClientHeight    =   1280
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5400
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' �L�����Z���{�^��
    ' �ʐ^�̈ړ����y�[�W�̃Z����I��
    Worksheets(SwapSourceSheet).Select
    Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
    ' �ʐ^�̈ړ����ƈړ���y�[�W�ԍ����N���A
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �ʐ^�𖖔��ֈړ��{�^��
    ' ���݂̃V�[�g�����擾
    SwapDestSheet = ActiveSheet.Name
    ' �ʐ^�𖖔��Ɉړ�������
    Call MoveToEnd
    ' �ʐ^�̈ړ����ƈړ���y�[�W�ԍ����N���A
    SwapSourceNo = 0
    SwapDestNo = 0
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏�������
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �{�^���𖳌��ɂ���
        CommandButton2.Enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ���[�U�[�t�H�[���̏I���O�̏���
    ' �u�~�v�ŕ��悤�Ƃ����ꍇ
    If CloseMode = vbFormControlMenu Then
        ' �ʐ^�̈ړ����y�[�W�̃Z����I��
        Worksheets(SwapSourceSheet).Select
        Range(PictureColumn(SwapSourceNo) & Format(PictureRow(SwapSourceNo))).MergeArea.Select
        ' �ʐ^�̈ړ����ƈړ���y�[�W�ԍ����N���A
        SwapSourceNo = 0
        SwapDestNo = 0
    End If
End Sub
