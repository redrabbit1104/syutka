Attribute VB_Name = "Module1"

'�o�׃V�[�g�̏o�̓{�^�����N���b�N�������̏���
Sub �o�׈˗����̏o�̓{�^��_Click()
    
'�v�����^�[��ʂ��Ăяo��
    Worksheets("�o�׈˗���").PrintPreview EnableChanges:=Ture
    
End Sub
'�o�׃V�[�g�̏o�͍s�I���{�^�����N���b�N�������̏���
Sub �o�͍s�I��_Click()

'���Z�v���Z�X���\��
With Application
     .Calculation = xlCalculationManual
     .EnableEvents = False
     .ScreenUpdating = False
End With

'�o�׃V�[�g�̍s�Ɨ�̕ϐ����`
Dim r As Long
selected_row = Selection.Row  '�o�׃V�[�g�őI�������Z���̍s�̒l��selected_row�ɑ��

'������

Worksheets("�o�׈˗���").Range("AB1").Value = ""
Worksheets("�o�׈˗���").Range("AD1").Value = ""
Worksheets("�o�׈˗���").Range("A3").Value = ""
Worksheets("�o�׈˗���").Range("H19").Value = ""
Worksheets("�o�׈˗���").Range("H22").Value = ""
Worksheets("�o�׈˗���").Range("K22").Value = ""
Worksheets("�o�׈˗���").Range("Z22").Value = ""
Worksheets("�o�׈˗���").Range("I40").Value = ""

For i = 1 To 12
    Worksheets("�o�׈˗���").Range("I" & (32 + i)).Value = ""
Next i

For i = 1 To 4
    Worksheets("�o�׈˗���").Range("H" & (22 + i * 2)).Value = ""
    Worksheets("�o�׈˗���").Range("K" & (22 + i * 2)).Value = ""
    Worksheets("�o�׈˗���").Range("Z" & (22 + i * 2)).Value = ""
Next i

'�o�׈˗����V�[�g�̊e���ځi���ʂ܂Łj�ɏo�׈˗����̃f�[�^���擾�����
Worksheets("�o�׈˗���").Range("AB1").Value = Worksheets("�o��").Range("E" & selected_row).Value
Worksheets("�o�׈˗���").Range("AD1").Value = Worksheets("�o��").Range("G" & selected_row).Value
Worksheets("�o�׈˗���").Range("A3").Value = Worksheets("�o��").Range("B" & selected_row).Value
Worksheets("�o�׈˗���").Range("H19").Value = Worksheets("�o��").Range("L" & selected_row).Value
Worksheets("�o�׈˗���").Range("H22").Value = Worksheets("�o��").Range("F" & selected_row).Value
Worksheets("�o�׈˗���").Range("K22").Value = Worksheets("�o��").Range("C" & selected_row).Value
Worksheets("�o�׈˗���").Range("Z22").Value = Worksheets("�o��").Range("M" & selected_row).Value
Worksheets("�o�׈˗���").Range("I40").Value = Worksheets("�o��").Range("N" & selected_row).Value

'�o�׃V�[�g�̗X�֔ԍ��`���t�於�̓��̓f�[�^��ϐ��ɑ���ipostcode�`recipient)
postcode = Worksheets("�o��").Range("I" & selected_row).Value
address1 = Worksheets("�o��").Range("J" & selected_row).Value
address2 = Worksheets("�o��").Range("K" & selected_row).Value

code = Worksheets("�o��").Range("F" & selected_row).Value
recipient = Worksheets("�o��").Range("H" & selected_row).Value
numbers = Worksheets("�o��").Range("M" & selected_row).Value
 
Dim x
x = 1
Do While Worksheets("�o��").Range("H" & (selected_row + x)).Value = "�㓯" Or Worksheets("�o��").Range("H" & (selected_row + x)).Value = ""
    Worksheets("�o�׈˗���").Range("H" & (22 + x * 2)).Value = Worksheets("�o��").Range("F" & (selected_row + x)).Value
    Worksheets("�o�׈˗���").Range("K" & (22 + x * 2)).Value = Worksheets("�o��").Range("C" & (selected_row + x)).Value
    Worksheets("�o�׈˗���").Range("Z" & (22 + x * 2)).Value = Worksheets("�o��").Range("M" & (selected_row + x)).Value
    If x = 4 Then Exit Do
    x = x + 1
Loop




'�X�֔ԍ����Ȃ��ꍇ�̏o�׈˗�������������ɂ�菈���𕪂���
'�X�֔ԍ�����ł͂Ȃ��ꍇ�̏���
If IsEmpty(postcode) = False Then
Worksheets("�o�׈˗���").Range("I33").Value = "��" & postcode
Worksheets("�o�׈˗���").Range("I34").Value = address1
Worksheets("�o�׈˗���").Range("I35").Value = address2
Worksheets("�o�׈˗���").Range("I36").Value = recipient
'���̑��̏ꍇ�̏����i�X�֔ԍ�����j
Else
Worksheets("�o�׈˗���").Range("I34").Value = recipient
Worksheets("�o�׈˗���").Range("I33").Value = ""
Worksheets("�o�׈˗���").Range("I36").Value = ""
Worksheets("�o�׈˗���").Range("I35").Value = ""
End If

'���Z���ʂ�\��
With Application
     .Calculation = xlCalculationAutomatic
     .EnableEvents = True
     .ScreenUpdating = True
End With

End Sub

