Attribute VB_Name = "Module6"
Sub �ٕ����ڂ��폜()

  Dim range As range
  Dim cell As range
  Dim chartSheet As Worksheet

  ' �V�[�g���u�O���t�v�̃V�[�g�ɂ���ٕ��̍��ږ����擾
  Set chartSheet = Worksheets("�O���t")
  chartSheet.Activate
  Set range = chartSheet.range(Cells(6, 2), Cells(6, 100))
  Set range = range.SpecialCells(xlCellTypeConstants)
  
  ' �폜���ڂ�I�����邽�߂̃��[�U�[�t�H�[�����쐬���ĕ\��
  UserForm1.Label1.Caption = "�폜���������ڂ�I�����Ă�������"
  For Each cell In range
    UserForm1.ComboBox1.AddItem (cell.Value)
  Next cell
  UserForm1.CommandButton1.Caption = "�폜����"
  
  UserForm1.Show
  
End Sub
