Attribute VB_Name = "Module21"
Option Explicit
Dim num As Integer


Sub �e�X�g()
  Do
    ' �ǉ����������ڂ���͂���{�b�N�X�̕\��
    Dim newTitle As String
    newTitle = InputBox(" �ǉ����������ږ�����͂��ĉ����� ")
    
    ' ���ڒǉ��t�H�[���ŃL�����Z���{�^���������ꂽ�ꍇ�A�������I������
    If StrPtr(newTitle) = 0 Then
      Exit Sub
    
    ' �ǉ����鍀�ڂ���̎��A���ږ��̓��͂𑣂����b�Z�[�W��\��
    ElseIf Trim(newTitle) = "" Then
      MsgBox "�ǉ����������ږ�����͂��Ă�������"
    
    ' �{�b�N�X�ɍ��ږ������͂��ꂽ���݈̂ȉ��̏��������s
    ElseIf newTitle <> "" Then
      
      ' ���i�̃��[�N�V�[�g�ɂ݈̂ȉ��̏������J��Ԃ����s
      Dim ws As Worksheet
      
      For Each ws In ThisWorkbook.Worksheets
        ws.Select
        Select Case ws.Name
          Case "�ʐ^", "�R�}���h�{�^��"
            ' �����ΏۊO�Ȃ̂ŁA�����Ȃ�
            
          Case "�O���t"
          ' �O���t�V�[�g�ɐV�������ڂ�ǉ����鏈��
          Dim graphSheet As Worksheet
          Set graphSheet = Worksheets("�O���t")
          
          ' �����́u�ٕ��̎�ނƔ��������v�̃O���t���폜
          Dim chartObj As ChartObject
          For Each chartObj In graphSheet.ChartObjects
            chartObj.Select
            If chartObj.chart.chartTitle.Text = "�ٕ��̎�ނƔ�������" Then
              chartObj.Delete
            End If
          Next
  
          Dim n As Long
          For n = 1 To 100
  
            If IsEmpty(graphSheet.Cells(6, n).Value) Then
              graphSheet.Cells(6, n).Value = newTitle
              graphSheet.range(Cells(7, n - 1), Cells(7, n - 1)).AutoFill Destination:=graphSheet.range(Cells(7, n - 1), Cells(7, n)), Type:=xlFillDefault
  
              graphSheet.Cells(6, n).Borders(xlEdgeTop).Weight = xlThin
              graphSheet.Cells(6, n).Borders(xlEdgeRight).Weight = xlThin
              graphSheet.Cells(6, n).Borders(xlEdgeBottom).LineStyle = xlDouble
  
              graphSheet.Cells(6, n).HorizontalAlignment = xlCenter
              graphSheet.Cells(6, n).VerticalAlignment = xlCenter
  
              Dim stringLength As Integer
              stringLength = Len(newTitle)
              If stringLength > 6 Then
                graphSheet.Columns(n).AutoFit
              End If
              
              ' �ٕ��̎�ނƔ��������̃O���t�������쐬
              ' �V���Ɂu�ٕ��̎�ނƔ��������v�̃O���t���쐬
              With graphSheet.Shapes.AddChart2.chart
                .HasTitle = True
                .chartTitle.Text = "�ٕ��̎�ނƔ�������"
                .ChartType = xlColumnClustered
                .SetSourceData range(Cells(6, "B"), Cells(7, n))
                
                ' �c���̃��x����\��
                With .Axes(xlValue)
                      .HasTitle = True
                      .AxisTitle.Text = "��������"
                      .AxisTitle.Orientation = xlVertical
                End With
                
                ' �O���t�̕\���ʒu�ƃT�C�Y�̐ݒ�
                With ActiveSheet.ChartObjects
                      .Top = range("A10").Top
                      .Left = range("A10").Left
                      .Height = 300
                      .Width = range(Cells(6, "B"), Cells(7, n)).Width
                End With
              End With
              Exit For
            End If
          Next n
              
          Case Else
            ' ���i�̃V�[�g�ɂ݈̂ȉ��̏��������s
            Dim i As Long
            For i = 1 To 100
            
              ' ���ڂ�ǉ����鏈��
              If IsEmpty(Cells(6, i).Value) Then
              
                '  �V�K���ڂ̗��ǉ����鏈��
                Cells(6, i).Value = newTitle
                range(Cells(7, i - 1), Cells(35, i - 1)).AutoFill Destination:=range(Cells(7, i - 1), Cells(35, i)), Type:=xlFillDefault
                
                '  �e�[�u���̘g���̎w��
                range(Cells(7, i), Cells(35, i)).Borders(xlEdgeLeft).Weight = xlThin
                Cells(6, i).Borders(xlEdgeTop).Weight = xlMedium
                Cells(6, i).Borders(xlEdgeRight).Weight = xlMedium
                Cells(6, i).Borders(xlEdgeLeft).Weight = xlThin
                
                '  �V�K�ǂ��������ڂ̃Z�����̔z�u�̎w��
                Cells(6, i).HorizontalAlignment = xlCenter
                Cells(6, i).VerticalAlignment = xlCenter
                
                ' �V�K�ǉ�������̕��𒲐�����
                Dim length As Integer
                length = Len(newTitle)
        
                If length > 6 Then
                  Columns(i).AutoFit
                End If
                
                ' �V�K�ǉ�������̃f�[�^���͔͈͓��̃f�[�^����ɂ��鏈���i�I�[�g�t�B���ŗׂ̃f�[�^���R�s�[���邽�߁j
                range(Cells(8, i), Cells(35, i)).Value = ""
                
                Exit For
              End If
              
            Next i
        End Select
      Next ws
      Exit Do
    End If
  Loop
End Sub




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



