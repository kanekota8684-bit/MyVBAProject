VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "������I�����Ă��������B"
   ClientHeight    =   6610
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   4572
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ���X�g�̍s���̍ő�l��萔�Ƃ��Đݒ肷��
Const myMaxRowCount As Long = 65536
' ���X�g�̃V�[�g����萔�Ƃ��Đݒ肷��
Const myItemSheetName As String = "�p��W"

Private Sub CheckBox2_Click()
' �܂�Ԃ��ĕ\���`�F�b�N�{�b�N�X
    If CheckBox2.Value Then
        CheckBox1.Enabled = False
    Else
        CheckBox1.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
' �R�s�[�{�^��
    Dim myTopRow As Long, myRowsCount As Long
    ' �I��͈͂̂����P��ڂ�����I������
    Selection.Columns(1).Select
    ' �I��͈͂̊J�n�s�����߂�
    myTopRow = Selection.Row
    ' �I��͈͂̍s�������߂�
    myRowsCount = Selection.Rows.Count
    ' �I��͈͂̍s�����K��l�𒴂���ꍇ
    If (myTopRow Mod 11) + myRowsCount > 10 Then
        ' �I��͈͂̍s���𐧌�����
        myRowsCount = 11 - (myTopRow Mod 11)
    End If
    ' �I��͈͂𐧌�����
    Selection.Resize(myRowsCount).Select
    ' �I��͈͂��R�s�[
    Selection.Copy
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton10_Click()
' �y�[�W�ǉ��{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' �y�[�W�ǉ����Ăяo���i�����͒ǉ������j
    Call AddPageProc(1)
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton11_Click()
' ����{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' ����v���r���[���Ăяo��
    UserForm1.Show vbModal
    ' ���[�U�[�t�H�[�����������
    Unload Me
End Sub

Private Sub CommandButton12_Click()
' �����Ɏ捞�{�^��
    ' ���[�U�[�t�H�[�����\���ɂ���
    Me.Hide
    ' �ꊇ��荞�݂��Ăяo��
    Call GetMultiPicture
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton13_Click()
' �K�p�{�^��
    Dim myTopRow As Long, myRowsCount As Long
    ' �I��͈͂̂����P��ڂ�����I������
    Selection.Columns(1).Select
    ' �I��͈͂̊J�n�s�����߂�
    myTopRow = Selection.Row
    ' �I��͈͂̍s�������߂�
    myRowsCount = Selection.Rows.Count
    ' �I��͈͂̍s�����K��l�𒴂���ꍇ
    If (myTopRow Mod 11) + myRowsCount > 10 Then
        ' �I��͈͂̍s���𐧌�����
        myRowsCount = 11 - (myTopRow Mod 11)
    End If
    ' �I��͈͂𐧌�����
    Selection.Resize(myRowsCount).Select
    ' �k�����đS�̂�\��
    Selection.ShrinkToFit = CheckBox1.Value
    ' �܂�Ԃ��đS�̂�\��
    Selection.WrapText = CheckBox2.Value
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' �\��t���{�^��
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �r���ȊO��\��t��
    Selection.PasteSpecial Paste:=xlPasteAllExceptBorders
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton3_Click()
' �L�����Z���{�^��
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton14_Click()
' �R�}���ɔԍ��U�����{�^��
    Me.Hide
    ' �R�}���̔ԍ��t�ԏ������Ăяo��
    Call SerialNumbering
    Unload Me
End Sub

Private Sub CommandButton15_Click()
' �ʐ^���ɔԍ��U�����{�^��
    Me.Hide
    ' �ʐ^���̔ԍ��t�ԏ������Ăяo��
    Call PictureNumbering
    Unload Me
End Sub

Private Sub CommandButton16_Click()
' �V���[�g�J�b�g�\���{�^��
    ' �V���[�g�J�b�g�\���t���O���Z�b�g
    ShortCutFlag = False
    Unload Me
End Sub

Private Sub CommandButton17_Click()
' �}�N���Ȃ��ŕۑ��I���{�^��
    Me.Hide
    ' �}�N���Ȃ��ŕۑ��I���������Ăяo��
    Call SaveWOMacro
    Unload Me
End Sub

Private Sub CommandButton18_Click()
' �B�e���\���{�^��
    Me.Hide
    ' �B�e�����\���V�[�P���X���Ăяo��
    Call PictureDateDispSequence
    Unload Me
End Sub

Private Sub CommandButton19_Click()
' �����̃y�[�W�폜�{�^��
    Me.Hide
    ' �����̃y�[�W�폜�������Ăяo��
    Call DeleteLastPages
    Unload Me
End Sub


Private Sub CommandButton4_Click()
' �Z���ɑ}���{�^���i�P�j
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �R���{�{�b�N�X�̃��X�g���I������Ă���ꍇ
    If ComboBox1.ListIndex <> -1 Then
        ' �R���{�{�b�N�X�̃��X�g�̒l���Z���ɑ}��
        Selection.Value = ComboBox1.List(ComboBox1.ListIndex)
        ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
        ComboList1 = ComboBox1.ListIndex
    ' �R���{�{�b�N�X�ɓ��͂��ꂽ�l�����X�g�ɖ����ꍇ
    ElseIf ComboBox1.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�����
        With Worksheets(myItemSheetName)
            ' ���X�g�̂P�s�ڂ��l�Ȃ��̏ꍇ
            If .Range("A1").Value = "" Then
                ' ���X�g�̂P�s�ڂɃR���{�{�b�N�X�̒l��ǉ�
                .Range("A1").Value = ComboBox1.Value
                ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
                ComboList1 = 0
            ' ���X�g�̂Q�s�ڂ��l�Ȃ��̏ꍇ
            ElseIf .Range("A2").Value = "" Then
                ' ���X�g�̂Q�s�ڂɃR���{�{�b�N�X�̒l��ǉ�
                .Range("A2").Value = ComboBox1.Value
                ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
                ComboList1 = 1
            ' ���X�g�̍ŏI�s���ő�l�����̏ꍇ
            ElseIf .Range("A1").End(xlDown).Row < myMaxRowCount Then
                ' ���X�g�̍ŏI�s�̉��ɃR���{�{�b�N�X�̒l��ǉ�
                .Range("A1").End(xlDown).Offset(1).Value = ComboBox1.Value
                ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
                ComboList1 = .Range("A1").End(xlDown).Row - 1
            End If
        End With
        ' �R���{�{�b�N�X�̒l���Z���ɑ}��
        Selection.Value = ComboBox1.Value
    End If
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton5_Click()
' �Z���ɑ}���{�^���i�Q�j
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �R���{�{�b�N�X�̃��X�g���I������Ă���ꍇ
    If ComboBox2.ListIndex <> -1 Then
        ' �R���{�{�b�N�X�̃��X�g�̒l���Z���ɑ}��
        Selection.Value = ComboBox2.List(ComboBox2.ListIndex)
        ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
        ComboList2 = ComboBox2.ListIndex
    ' �R���{�{�b�N�X�ɓ��͂��ꂽ�l�����X�g�ɂȂ��ꍇ
    ElseIf ComboBox2.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�����
        With Worksheets(myItemSheetName)
            If .Range("B1").Value = "" Then
                .Range("B1").Value = ComboBox2.Value
                ComboList2 = 0
            ElseIf .Range("B2").Value = "" Then
                .Range("B2").Value = ComboBox2.Value
                ComboList2 = 1
            ElseIf .Range("B2").End(xlDown).Row < myMaxRowCount Then
                .Range("B1").End(xlDown).Offset(1).Value = ComboBox2.Value
                ComboList2 = .Range("B1").End(xlDown).Row - 1
            End If
        End With
        ' �R���{�{�b�N�X�̒l���Z���ɑ}��
        Selection.Value = ComboBox2.Value
    End If
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton6_Click()
' �Z���ɑ}���{�^���i�R�j
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �R���{�{�b�N�X�̃��X�g���I������Ă���ꍇ
    If ComboBox3.ListIndex <> -1 Then
        ' �R���{�{�b�N�X�̃��X�g�̒l���Z���ɑ}��
        Selection.Value = ComboBox3.List(ComboBox3.ListIndex)
        ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
        ComboList3 = ComboBox3.ListIndex
    ' �R���{�{�b�N�X�ɓ��͂��ꂽ�l�����X�g�ɂȂ��ꍇ
    ElseIf ComboBox3.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�����
        With Worksheets(myItemSheetName)
            If .Range("C1").Value = "" Then
                .Range("C1").Value = ComboBox3.Value
                ComboList3 = 0
            ElseIf .Range("C2").Value = "" Then
                .Range("C2").Value = ComboBox3.Value
                ComboList3 = 1
            ElseIf .Range("C1").End(xlDown).Row < myMaxRowCount Then
                .Range("C1").End(xlDown).Offset(1).Value = ComboBox3.Value
                ComboList3 = .Range("C1").End(xlDown).Row - 1
            End If
        End With
        ' �R���{�{�b�N�X�̒l���Z���ɑ}��
        Selection.Value = ComboBox3.Value
    End If
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton7_Click()
' �Z���ɑ}���{�^���i�S�j
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �R���{�{�b�N�X�̃��X�g���I������Ă���ꍇ
    If ComboBox4.ListIndex <> -1 Then
        ' �R���{�{�b�N�X�̃��X�g�̒l���Z���ɑ}��
        Selection.Value = ComboBox4.List(ComboBox4.ListIndex)
        ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
        ComboList4 = ComboBox4.ListIndex
    ' �R���{�{�b�N�X�ɓ��͂��ꂽ�l�����X�g�ɂȂ��ꍇ
    ElseIf ComboBox4.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�����
        With Worksheets(myItemSheetName)
            If .Range("D1").Value = "" Then
                .Range("D1").Value = ComboBox4.Value
                ComboList4 = 0
            ElseIf .Range("D2").Value = "" Then
                .Range("D2").Value = ComboBox4.Value
                ComboList4 = 1
            ElseIf .Range("D1").End(xlDown).Row < myMaxRowCount Then
                .Range("D1").End(xlDown).Offset(1).Value = ComboBox4.Value
                ComboList4 = .Range("D1").End(xlDown).Row - 1
            End If
        End With
        ' �R���{�{�b�N�X�̒l���Z���ɑ}��
        Selection.Value = ComboBox4.Value
    End If
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton8_Click()
' �Z���ɑ}���{�^���i�T�j
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �R���{�{�b�N�X�̃��X�g���I������Ă���ꍇ
    If ComboBox5.ListIndex <> -1 Then
        ' �R���{�{�b�N�X�̃��X�g�̒l���Z���ɑ}��
        Selection.Value = ComboBox5.List(ComboBox5.ListIndex)
        ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
        ComboList5 = ComboBox5.ListIndex
    ' �R���{�{�b�N�X�ɓ��͂��ꂽ�l�����X�g�ɂȂ��ꍇ
    ElseIf ComboBox5.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�����
        With Worksheets(myItemSheetName)
            If .Range("E1").Value = "" Then
                .Range("E1").Value = ComboBox5.Value
                ComboList5 = 0
            ElseIf .Range("E2").Value = "" Then
                .Range("E2").Value = ComboBox5.Value
                ComboList5 = 1
            ElseIf .Range("E1").End(xlDown).Row < myMaxRowCount Then
                .Range("E1").End(xlDown).Offset(1).Value = ComboBox5.Value
                ComboList5 = .Range("E1").End(xlDown).Row - 1
            End If
        End With
        ' �R���{�{�b�N�X�̒l���Z���ɑ}��
        Selection.Value = ComboBox5.Value
    End If
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub CommandButton9_Click()
' �Z���ɑ}���{�^���i�U�j
    ' �G���[�����������玟�̏�����
    On Error Resume Next
    ' �R���{�{�b�N�X�̃��X�g���I������Ă���ꍇ
    If ComboBox6.ListIndex <> -1 Then
        ' �R���{�{�b�N�X�̃��X�g�̒l���Z���ɑ}��
        Selection.Value = ComboBox6.List(ComboBox6.ListIndex)
        ' �R���{�{�b�N�X�̐擪�ʒu��ۑ�
        ComboList6 = ComboBox6.ListIndex
    ' �R���{�{�b�N�X�ɓ��͂��ꂽ�l�����X�g�ɂȂ��ꍇ
    ElseIf ComboBox6.MatchFound = False Then
        ' ���X�g�ɒl��ǉ�����
        With Worksheets(myItemSheetName)
            If .Range("F1").Value = "" Then
                .Range("F1").Value = ComboBox6.Value
                ComboList6 = 0
            ElseIf .Range("F2").Value = "" Then
                .Range("F2").Value = ComboBox6.Value
                ComboList6 = 1
            ElseIf .Range("F1").End(xlDown).Row < myMaxRowCount Then
                .Range("F1").End(xlDown).Offset(1).Value = ComboBox6.Value
                ComboList6 = .Range("F1").End(xlDown).Row - 1
            End If
        End With
        ' �R���{�{�b�N�X�̒l���Z���ɑ}��
        Selection.Value = ComboBox6.Value
    End If
    ' ���[�U�[�t�H�[���̉��
    Unload Me
End Sub

Private Sub UserForm_Initialize()
' ���[�U�[�t�H�[���̏����ݒ�
    Dim myWorksheet As Worksheet
    Dim mySheetExist As Boolean
    ' �V���[�g�J�b�g�\���t���O�����Z�b�g
    ShortCutFlag = True
    ' �`�F�b�N�{�b�N�X�̏����ݒ�
    CheckBox1.Value = True
    CheckBox2.Value = False
    ' �V�[�g���ی삳��Ă���ꍇ
    If ActiveSheet.ProtectContents Then
        ' �R�}���h�{�^���𖳌��ɂ���
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        CommandButton4.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton9.Enabled = False
        CommandButton10.Enabled = False
        CommandButton12.Enabled = False
        CommandButton13.Enabled = False
        CommandButton14.Enabled = False
        CommandButton15.Enabled = False
        CommandButton17.Enabled = False
        CommandButton18.Enabled = False
        CommandButton19.Enabled = False
        ' �R���{�{�b�N�X�P����U�𖳌��ɂ���
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False
        ' �����ݒ���I��
        Exit Sub
    End If
    ' �N���b�v�{�[�h���e�L�X�g�łȂ��ꍇ
    If Application.ClipboardFormats(1) <> xlClipboardFormatText Then
        ' �\��t���{�^���𖳌��ɂ���
        CommandButton2.Enabled = False
    End If
    ' �B�e���\���{�^������ݒ�
    If Range(PictureDateFlag).Value <> 0 Then
        CommandButton18.Caption = "�B�e���\���n�e�e(F)"
    Else
        CommandButton18.Caption = "�B�e���\���n�m(F)"
    End If
    ' ���[�N�V�[�g�t���O�����Z�b�g
    mySheetExist = False
    For Each myWorksheet In Worksheets
        ' ���[�N�V�[�g(myItemSheetName)�����݂����
        If myWorksheet.Name = myItemSheetName Then
            ' ���[�N�V�[�g�t���O���Z�b�g
            mySheetExist = True
            ' ���肩���������𔲂��o��
            Exit For
        End If
    Next
    ' ���[�N�V�[�g(myItenSheetName)���������
    If mySheetExist = False Then
        ' �R�}���h�{�^���S����X�𖳌��ɂ���
        CommandButton4.Enabled = False
        CommandButton5.Enabled = False
        CommandButton6.Enabled = False
        CommandButton7.Enabled = False
        CommandButton8.Enabled = False
        CommandButton9.Enabled = False
        ' �R���{�{�b�N�X�P����U�𖳌��ɂ���
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False
        ' �����ݒ���I��
        Exit Sub
    End If
    With Worksheets(myItemSheetName)
        ' �R���{�{�b�N�X�P�̃Z���͈͂��X�V
        If .Range("A1") = "" Or .Range("A2") = "" Then
            ComboBox1.RowSource = .Range("A1").Address(External:=True)
        Else
            ComboBox1.RowSource = _
                    .Range("A1:" & "A" & Format(.Range("A1").End(xlDown).Row)).Address(External:=True)
        End If
        ' �R���{�{�b�N�X�P�̐擪�ʒu�𕜌�
        ComboBox1.ListIndex = ComboList1
        ' �R���{�{�b�N�X�Q�̃Z���͈͂��X�V
        If .Range("B1") = "" Or .Range("B2") = "" Then
            ComboBox2.RowSource = .Range("B1").Address(External:=True)
        Else
            ComboBox2.RowSource = _
                    .Range("B1:" & "B" & Format(.Range("B1").End(xlDown).Row)).Address(External:=True)
        End If
        ' �R���{�{�b�N�X�Q�̐擪�ʒu�𕜌�
        ComboBox2.ListIndex = ComboList2
        ' �R���{�{�b�N�X�R�̃Z���͈͂��X�V
        If .Range("C1") = "" Or .Range("C2") = "" Then
            ComboBox3.RowSource = .Range("C1").Address(External:=True)
        Else
            ComboBox3.RowSource = _
                    .Range("C1:" & "C" & Format(.Range("C1").End(xlDown).Row)).Address(External:=True)
        End If
        ' �R���{�{�b�N�X�R�̐擪�ʒu�𕜌�
        ComboBox3.ListIndex = ComboList3
        ' �R���{�{�b�N�X�S�̃Z���͈͂��X�V
        If .Range("D1") = "" Or .Range("D2") = "" Then
            ComboBox4.RowSource = .Range("D1").Address(External:=True)
        Else
            ComboBox4.RowSource = _
                    .Range("D1:" & "D" & Format(.Range("D1").End(xlDown).Row)).Address(External:=True)
        End If
        ' �R���{�{�b�N�X�S�̐擪�ʒu�𕜌�
        ComboBox4.ListIndex = ComboList4
        ' �R���{�{�b�N�X�T�̃Z���͈͂��X�V
        If .Range("E1") = "" Or .Range("E2") = "" Then
            ComboBox5.RowSource = .Range("E1").Address(External:=True)
        Else
            ComboBox5.RowSource = _
                    .Range("E1:" & "E" & Format(.Range("E1").End(xlDown).Row)).Address(External:=True)
        End If
        ' �R���{�{�b�N�X�T�̐擪�ʒu�𕜌�
        ComboBox5.ListIndex = ComboList5
        ' �R���{�{�b�N�X�U�̃Z���͈͂��X�V
        If .Range("F1") = "" Or .Range("F2") = "" Then
            ComboBox6.RowSource = .Range("F1").Address(External:=True)
        Else
            ComboBox6.RowSource = _
                    .Range("F1:" & "F" & Format(.Range("F1").End(xlDown).Row)).Address(External:=True)
        End If
        ' �R���{�{�b�N�X�U�̐擪�ʒu�𕜌�
        ComboBox6.ListIndex = ComboList6
    End With
End Sub
