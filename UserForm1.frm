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
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Select Case ws.Name
        Case "�ʐ^", "�R�}���h�{�^��"
          ' �����ΏۊO�Ȃ̂ŁA�����Ȃ�
          
        Case "�O���t"
'        ' �O���t�̃V�[�g�ɑ΂��Ĉȉ��̏��������s
'        Dim range As range
'        Dim cell As range

        Set chartSheet = Worksheets("�O���t")

'        ' �O���t�V�[�g���ٕ̈����ڂ̎擾
'        Set range = chartSheet.range(Cells(6, 2), Cells(6, 100))
'        Set range = range.SpecialCells(xlCellTypeConstants)
'
'        For Each cell In range
'          ' �O���t�V�[�g���ٕ̈����ږ��ƃ��[�U�[�t�H�[���őI�������ٕ����ږ�����v�������A�ȉ��̏��������s
'          If cell.Value = UserForm1.ComboBox1.Value Then
'            cell.ClearContents
'
'            Dim i As Long
'            For i = 2 To 100
'              If IsEmpty(Cells(6, i + 1).Value) Then
'                Exit For
'              End If
'              If IsEmpty(Cells(6, i).Value) Then
'                Cells(6, i + 1).Offset(0, -1).Value = Cells(6, i + 1).Value
'              End If
'            Next i
'            Exit For
'          End If
'        Next cell
      Dim i As Long
      For i = 2 To 100
        If IsEmpty(Cells(6, i + 1).Value) Then
          Exit For
        End If
        If Cells(6, i).Value = UserForm1.ComboBox1.Value Then
          Cells(6, i).ClearContents
          Cells(7, i).ClearContents
        End If
        If IsEmpty(Cells(6, i).Value) Then
          Cells(6, i + 1).Offset(0, -1).Value = Cells(6, i + 1).Value
          Cells(6, i + 1).ClearContents
          Cells(7, i + 1).Offset(0, -1).Value = Cells(7, i + 1).Value
          Cells(7, i + 1).ClearContents
        End If
      Next i
        
      End Select
    Next ws
    chartSheet.Select
  End If
End Sub
