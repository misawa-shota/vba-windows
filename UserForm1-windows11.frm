VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' �폜�{�^���������ꂽ���ȉ��̏��������s
Private Sub CommandButton1_Click()
  ' �폜���������ږ������I���̎����b�Z�[�W��\�����č��ڑI����ʂɖ߂�
  If UserForm1.ComboBox1.Value = "" Then
    MsgBox "�폜���������ڂ�I�����Ă�������"
    Exit Sub
  End If
  
  ' �I�����ꂽ���ڂ�{���ɍ폜���邩�m�F���郁�b�Z�[�W��\��
  Dim button As Integer
  button = MsgBox("���ږ��u" & UserForm1.ComboBox1.Value & "�v���폜���܂���?", vbYesNo + vbQuestion + vbDefaultButton2, "���ڍ폜")
  ' �u�������v��I���������A�������I�����č��ڑI����ʂɖ߂�
  If button = vbNo Then
    Exit Sub
  
  ' �u�͂��v���I�����ꂽ���A�ȉ��̏��������s
  ElseIf button = vbYes Then
    Dim ws As Worksheet
    Dim chartSheet As Worksheet
    Set chartSheet = Worksheets("�O���t")
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Select Case ws.Name
        Case "�ʐ^", "�R�}���h�{�^��"
          ' �����ΏۊO�Ȃ̂ŁA�����Ȃ�
          
        Case "�O���t"
        ' �O���t�V�[�g�Ɉȉ��̏��������s
        Dim i As Long
        For i = 2 To 100
          ' i + 1�Ԗڂ̒l����̎��Ai�Ԗڂ̗����폜���āA�J��Ԃ��������I������
          If IsEmpty(Cells(6, i + 1).Value) Then
            range(Cells(6, i), Cells(7, i)).Delete
            Exit For
          End If
          ' i�Ԗڂ̃Z���̒l�ƃ��[�U�[�t�H�[���őI�����ꂽ�l����v�������Ai�Ԗڂ̃Z���̒l����ɂ���
          If Cells(6, i).Value = UserForm1.ComboBox1.Value Then
            Cells(6, i).ClearContents
          End If
          ' i�Ԗڂ̃Z���̒l����̎��Ai + 1�Ԗڂ̒l��i�ԖڂɈړ����āAi + 1�Ԗڂ̒l����ɂ���
          If IsEmpty(Cells(6, i).Value) Then
             Cells(6, i + 1).Offset(0, -1).Value = Cells(6, i + 1).Value
             Cells(6, i + 1).ClearContents
          End If
        Next i
        
        Case Else
        ' ���i�̃V�[�g�Ɉȉ��̏��������s
        Dim n As Long
        For n = 4 To 100
          ' ���[�U�[�t�H�[���őI�����ꂽ�ٕ����ږ��Ɉ�v���鍀�ڂ��폜���āA����ȍ~�̗������Ɉړ�������
          If Cells(6, n).Value = UserForm1.ComboBox1.Value Then
            range(Cells(6, n), Cells(35, n)).Delete Shift:=xlToLeft
            ' �폜����ٕ����ڂ��Ō�̍��ڂ̎��A���̗��̍��̘g���𑾂����鏈��
            If IsEmpty(Cells(6, n + 1).Value) Then
              range(Cells(6, n), Cells(35, n)).Borders(xlEdgeLeft).Weight = xlMedium
            End If
            Exit For
          End If
        Next n
        
      End Select
    Next ws
    chartSheet.Select
  End If
  Unload Me
End Sub
