Attribute VB_Name = "Module1"
Option Explicit
Sub ��A�N�e�B�u�V�[�g�ւ̑���G���[��()
    Dim �z��()
    Sheet1.Activate
    �z�� = Range(Cells(1, 1), Cells(5, 5))
    Sheet2.Range(Cells(1, 1), Cells(5, 5)) = �z��
End Sub
Sub OS���_���猩���G���[��()
    Dim �z��()
    Sheet1.Activate
    �z�� = Range(Cells(1, 1), Cells(5, 5))
    'Cells�̐e�I�u�W�F�N�g���ȗ�����Ă��� �� ActiveSheet(= Sheet1)���e�I�u�W�F�N�g���Ɛ���
    Sheet2.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(5, 5)) = �z��
End Sub
Sub OK��1()
    Dim �z��()
    Sheet1.Activate
    �z�� = Range(Cells(1, 1), Cells(5, 5))
    Range(Sheet2.Cells(1, 1), Sheet2.Cells(5, 5)) = �z��
    'Range�̐e�I�u�W�F�N�g���ȗ�����Ă���AActiveSheet(= Sheet1)���e�I�u�W�F�N�g���Ɛ��肳�ꂻ���ɂ������邪�A
    '����q�\���͓������珇�ɓ��肳��Ă���(Cells���� �� Cells������q���ꂽRange����̏�)���߁A
    '��薳���͗l�B(Sheet2.Cells�̓���q�ɂ���ĕ\�������Range �� Sheet2���e�I�u�W�F�N�g���Ɛ���)
    '�����ۂ̃R�[�f�B���O�ł�With�\���ɂ��e�I�u�W�F�N�g�L�q�ȗ����g���Ɗy�ŉǐ����オ��܂��B
    End With
End Sub
Sub OK��2()
    Dim �z��()
    Sheet1.Activate
    �z�� = Range(Cells(1, 1), Cells(5, 5))
    Sheet2.Cells(1, 1).Resize(5, 5) = �z��
End Sub
