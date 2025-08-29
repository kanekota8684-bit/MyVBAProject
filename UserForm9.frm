VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "�B�e���̕ҏW"
   ClientHeight    =   1090
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3876
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myPictureNo As Integer, myPictureName As String

Private Sub CommandButton1_Click()
' ����{�^��
    Dim myDate As String, myType As Integer
    myDate = Format(TextBox1.Value) & "/" & Format(TextBox2.Value) & "/" & Format(TextBox3.Value) _
        & " " & Format(TextBox4.Value) & ":" & Format(TextBox5.Value)
    ' �l�̃`�F�b�N
    If Not IsDate(myDate) Then
        MsgBox "�l���s���ł��B", vbOKOnly + vbExclamation, "���m�点"
        Exit Sub
    End If
    ' �l���i�[
    Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value = myDate
    ' �B�e���̕\���`���t���O���G���R�[�h
    myType = Range(PictureDateType).Value + Range(DateSeparator).Value * 2 + Range(WeekDisp).Value * 4 + Range(WeekLang).Value * 8
    ' �e�L�X�g�{�b�N�X�̕�������������
    With ActiveSheet.Shapes(Application.Caller)
        .TextFrame.Characters.Text = PictureDateFormat(myDate, myType)
    End With
    ' �e�L�X�g�{�b�N�X�̈ʒu����
    Call PictureDatePosition(myPictureName)
    ' ���[�U�[�t�H�[�������
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �L�����Z���{�^��
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏���������
    Dim myPictureDate As String
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �e�R���g���[���𖳌��ɂ���
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        CommandButton1.Enabled = False
        Exit Sub
    End If
    ' �ʐ^�̃y�[�W�ԍ������߂�
    myPictureNo = GetPictureNoFromDate(Application.Caller)
    ' �ʐ^�̎B�e���f�[�^���擾
    myPictureDate = PictureDateFormat(Range(PictureDateBuffer & Format(MinDataBuffer + myPictureNo)).Value, 0)
    ' �ʐ^�̖��O���擾
    myPictureName = Range(PictureNameBuffer & Format(MinDataBuffer + myPictureNo)).Value
    ' �B�e���f�[�^�����t�f�[�^�Ȃ�
    If IsDate(myPictureDate) Then
        ' �e�L�X�g�{�b�N�X�̏����l��ݒ�
        TextBox1.Value = Format(myPictureDate, "yyyy")
        TextBox2.Value = Format(myPictureDate, "m")
        TextBox3.Value = Format(myPictureDate, "d")
        TextBox4.Value = Format(myPictureDate, "h")
        TextBox5.Value = Format(myPictureDate, "n")
    ' ���t�f�[�^�Ƃ��Ĉ����Ȃ��ꍇ
    Else
        TextBox1.Value = "****"
        TextBox2.Value = "**"
        TextBox3.Value = "**"
        TextBox4.Value = "**"
        TextBox5.Value = "**"
    End If
End Sub
