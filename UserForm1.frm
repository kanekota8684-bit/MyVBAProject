VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�������O�̏�����I�����Ă��������B"
   ClientHeight    =   1320
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
' ����{�^��
    Me.Hide
    Call PreviewPrint
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �L�����Z���{�^��
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' �R�}���ɔԍ��U�����{�^��
    Me.Hide
    Call SerialNumbering
    Call PreviewPrint
    Unload Me
End Sub

Private Sub CommandButton4_Click()
' �ʐ^���ɔԍ��U�����{�^��
    Me.Hide
    Call PictureNumbering
    Call PreviewPrint
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏����ݒ�
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �{�^���𖳌��ɂ���
        CommandButton3.Enabled = False
        CommandButton4.Enabled = False
    End If
End Sub
