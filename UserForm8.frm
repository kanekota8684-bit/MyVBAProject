VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "�B�e������\�����܂��B"
   ClientHeight    =   2895
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4968
   OleObjectBlob   =   "UserForm8.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox4_Click()
' �j���\���`�F�b�N�{�b�N�X
    ' �j���\���`�F�b�N�{�b�N�X���n�m�Ȃ�
    If CheckBox4.Value = True Then
        ' �j���\�����ꃉ�W�I�{�^����L���ɂ���
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    ' �j���\���`�F�b�N�{�b�N�X���n�e�e�Ȃ�
    Else
        ' �j���\�����ꃉ�W�I�{�^���𖳌��ɂ���
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
    End If
End Sub

Private Sub CommandButton1_Click()
' ����{�^��
    ' �t�H���g�T�C�Y�����l�łȂ��ꍇ
    If Not IsNumeric(ComboBox1.Value) Then
        MsgBox "�t�H���g�T�C�Y�́A���l����͂��Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Exit Sub
    ElseIf ComboBox1.Value < 0 Then
        MsgBox "�t�H���g�T�C�Y�́A�[���ȏ�̐��l����͂��Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Exit Sub
    End If
    ' �t�H���g�T�C�Y���L��
    Range(DateFontSize).Value = ComboBox1.Value
    ' �����ɂ���t���O���L��
    If CheckBox1.Value = True Then
        Range(DateFontBold).Value = 1
    Else
        Range(DateFontBold).Value = ""
    End If
    ' �B�e�����̎������\�������邩���ɂ�������\�������邩�̃t���O���L��
    If CheckBox2.Value = True Then
        Range(PictureDateType).Value = 1
    Else
        Range(PictureDateType).Value = ""
    End If
    ' �t�H���g�̐F�����l�Ŗ����ꍇ
    If (Not IsNumeric(TextBox1.Value)) Or (Not IsNumeric(TextBox2.Value)) Or (Not IsNumeric(TextBox3.Value)) Then
        MsgBox "�t�H���g�̐F�́A���l����͂��Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Exit Sub
    End If
    ' �t�H���g�̐F���O�`�Q�T�T�̒l�łȂ��ꍇ
    If TextBox1.Value < 0 Or TextBox1.Value > 255 Or _
        TextBox2.Value < 0 Or TextBox2.Value > 255 Or _
        TextBox3.Value < 0 Or TextBox3.Value > 255 Then
        MsgBox "�t�H���g�̐F�́A0 �` 255 �̐��l�Ŏw�肵�Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Exit Sub
    End If
    ' �t�H���g�̐F���L��
    Range(DateFontColorR).Value = Int(TextBox1.Value)
    Range(DateFontColorG).Value = Int(TextBox2.Value)
    Range(DateFontColorB).Value = Int(TextBox3.Value)
    ' �e�L�X�g�{�b�N�X�̕\���ʒu�I�t�Z�b�g�����l�łȂ��ꍇ
    If (Not IsNumeric(TextBox4.Value)) Or (Not IsNumeric(TextBox5.Value)) Then
        MsgBox "�\���ʒu�̃I�t�Z�b�g�́A���l����͂��Ă��������B", vbOKOnly + vbExclamation, "���m�点"
        Exit Sub
    End If
    ' �e�L�X�g�{�b�N�X�̕\���ʒu�I�t�Z�b�g�̒l���L��
    Range(DateXOffset).Value = TextBox4.Value / XUnit
    Range(DateYOffset).Value = TextBox5.Value / YUnit
    ' �B�e����؂�L���u�D�v�̎g�p�t���O���L��
    If CheckBox3.Value = True Then
        Range(DateSeparator).Value = 1
    Else
        Range(DateSeparator).Value = ""
    End If
    ' �j���\���̒l���L��
    If CheckBox4.Value Then
        Range(WeekDisp).Value = 1
    Else
        Range(WeekDisp).Value = ""
    End If
    ' �j���\������̒l���L��
    If OptionButton2.Value Then
        Range(WeekLang).Value = 1
    Else
        Range(WeekLang).Value = ""
    End If
    ' ��ʕ\���̍X�V�����Ȃ��悤�ɂ���
    Application.ScreenUpdating = False
    ' �B�e���������������
    Call PictureDateOFF
    ' �B�e����\��������
    Call PictureDateON
    ' ��ʕ\���̍X�V��������
    Application.ScreenUpdating = True
    ' ���[�U�[�t�H�[�������
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �L�����Z���{�^��
    ' ���[�U�[�t�H�[�������
    Unload Me
End Sub

Private Sub TextBox1_Change()
' �e�L�X�g�{�b�N�X�̒l���ς�����ꍇ
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ���x���̐F��ݒ�
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
            ' ���[�U�[�t�H�[�����ĕ`��
            Me.Repaint
        End If
    End If
End Sub

Private Sub TextBox2_Change()
' �e�L�X�g�{�b�N�X�̒l���ς�����ꍇ
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ���x���̐F��ݒ�
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
            ' ���[�U�[�t�H�[�����ĕ`��
            Me.Repaint
        End If
    End If
End Sub

Private Sub TextBox3_Change()
' �e�L�X�g�{�b�N�X�̒l���ς�����ꍇ
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ���x���̐F��ݒ�
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
            ' ���[�U�[�t�H�[�����ĕ`��
            Me.Repaint
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏���������
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �e�R���g���[���𖳌��ɂ���
        CommandButton1.Enabled = False
        ComboBox1.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
        Exit Sub
    End If
    ' �R���{�{�b�N�X�̃��X�g���쐬
    With ComboBox1
        .AddItem 10, 0
        .AddItem 12, 1
        .AddItem 14, 2
        .AddItem 16, 3
    End With
    ' �R���{�{�b�N�X�̒l��ݒ�
    ComboBox1.Value = Range(DateFontSize).Value
    ' �R���{�{�b�N�X�̒l�����X�g�ɂȂ��ꍇ
    If ComboBox1.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�
        ComboBox1.AddItem Range(DateFontSize).Value, 4
    End If
    ' �`�F�b�N�{�b�N�X�̏����ݒ�i�����̐ݒ�j
    If Range(DateFontBold).Value <> 0 Then
        CheckBox1.Value = True
    Else
        CheckBox1.Value = False
    End If
    ' �e�L�X�g�{�b�N�X�̏����ݒ�i�t�H���g�̐F�j
    TextBox1.Value = Range(DateFontColorR).Value
    If TextBox1.Value = "" Then
        TextBox1.Value = 0
    End If
    TextBox2.Value = Range(DateFontColorG).Value
    If TextBox2.Value = "" Then
        TextBox2.Value = 0
    End If
    TextBox3.Value = Range(DateFontColorB).Value
    If TextBox3.Value = "" Then
        TextBox3.Value = 0
    End If
    ' ���x���̐F��ݒ�
    If IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) Then
        If Int(TextBox1.Value) >= 0 And Int(TextBox1.Value) <= 255 And _
            Int(TextBox2.Value) >= 0 And Int(TextBox2.Value) <= 255 And _
            Int(TextBox3.Value) >= 0 And Int(TextBox3.Value) <= 255 Then
            ' ���x���̐F��ݒ�
            Label7.ForeColor = RGB(Int(TextBox1.Value), Int(TextBox2.Value), Int(TextBox3.Value))
        End If
    End If
    ' �`�F�b�N�{�b�N�X�̏����ݒ�i�B�e���̂ݕ\���j
    If Range(PictureDateType).Value <> 0 Then
        CheckBox2.Value = True
    Else
        CheckBox2.Value = False
    End If
    ' �e�L�X�g�{�b�N�X�̏����ݒ�i�\���ʒu�I�t�Z�b�g�j
    TextBox4.Value = Round(Range(DateXOffset).Value * XUnit, 1)
    TextBox5.Value = Round(Range(DateYOffset).Value * YUnit, 1)
    ' �`�F�b�N�{�b�N�X�̏����ݒ�i�B�e����؂�̐ݒ�j
    If Range(DateSeparator).Value <> 0 Then
        CheckBox3.Value = True
    Else
        CheckBox3.Value = False
    End If
    ' �I�v�V�����{�^���̏����ݒ�i�j���\���̌���j
    If Range(WeekLang).Value <> 0 Then
        OptionButton2.Value = True
    Else
        OptionButton1.Value = True
    End If
    ' �`�F�b�N�{�b�N�X�̏����ݒ�i�j���\���̐ݒ�j
    If Range(WeekDisp).Value <> 0 Then
        CheckBox4.Value = True
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    Else
        CheckBox4.Value = False
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
    End If
End Sub

