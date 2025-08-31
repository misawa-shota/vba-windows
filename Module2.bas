Attribute VB_Name = "Module21"
Option Explicit
Dim num As Integer


Sub テスト()
  Do
    ' 追加したい項目を入力するボックスの表示
    Dim newTitle As String
    newTitle = InputBox(" 追加したい項目名を入力して下さい ")
    
    ' 項目追加フォームでキャンセルボタンが押された場合、処理を終了する
    If StrPtr(newTitle) = 0 Then
      Exit Sub
    
    ' 追加する項目が空の時、項目名の入力を促すメッセージを表示
    ElseIf Trim(newTitle) = "" Then
      MsgBox "追加したい項目名を入力してください"
    
    ' ボックスに項目名が入力された時のみ以下の処理を実行
    ElseIf newTitle <> "" Then
      
      ' 商品のワークシートにのみ以下の処理を繰り返し実行
      Dim ws As Worksheet
      
      For Each ws In ThisWorkbook.Worksheets
        ws.Select
        Select Case ws.Name
          Case "写真", "コマンドボタン"
            ' 処理対象外なので、処理なし
            
          Case "グラフ"
          ' グラフシートに新しい項目を追加する処理
          Dim graphSheet As Worksheet
          Set graphSheet = Worksheets("グラフ")
          
          ' 既存の「異物の種類と発生件数」のグラフを削除
          Dim chartObj As ChartObject
          For Each chartObj In graphSheet.ChartObjects
            chartObj.Select
            If chartObj.chart.chartTitle.Text = "異物の種類と発生件数" Then
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
              
              ' 異物の種類と発生件数のグラフを自動作成
              ' 新たに「異物の種類と発生件数」のグラフを作成
              With graphSheet.Shapes.AddChart2.chart
                .HasTitle = True
                .chartTitle.Text = "異物の種類と発生件数"
                .ChartType = xlColumnClustered
                .SetSourceData range(Cells(6, "B"), Cells(7, n))
                
                ' 縦軸のラベルを表示
                With .Axes(xlValue)
                      .HasTitle = True
                      .AxisTitle.Text = "発生件数"
                      .AxisTitle.Orientation = xlVertical
                End With
                
                ' グラフの表示位置とサイズの設定
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
            ' 商品のシートにのみ以下の処理を実行
            Dim i As Long
            For i = 1 To 100
            
              ' 項目を追加する処理
              If IsEmpty(Cells(6, i).Value) Then
              
                '  新規項目の列を追加する処理
                Cells(6, i).Value = newTitle
                range(Cells(7, i - 1), Cells(35, i - 1)).AutoFill Destination:=range(Cells(7, i - 1), Cells(35, i)), Type:=xlFillDefault
                
                '  テーブルの枠線の指定
                range(Cells(7, i), Cells(35, i)).Borders(xlEdgeLeft).Weight = xlThin
                Cells(6, i).Borders(xlEdgeTop).Weight = xlMedium
                Cells(6, i).Borders(xlEdgeRight).Weight = xlMedium
                Cells(6, i).Borders(xlEdgeLeft).Weight = xlThin
                
                '  新規追した加項目のセル内の配置の指定
                Cells(6, i).HorizontalAlignment = xlCenter
                Cells(6, i).VerticalAlignment = xlCenter
                
                ' 新規追加した列の幅を調整する
                Dim length As Integer
                length = Len(newTitle)
        
                If length > 6 Then
                  Columns(i).AutoFit
                End If
                
                ' 新規追加した列のデータ入力範囲内のデータを空にする処理（オートフィルで隣のデータをコピーするため）
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




Sub 異物項目を削除()

  Dim range As range
  Dim cell As range
  Dim chartSheet As Worksheet

  ' シート名「グラフ」のシートにある異物の項目名を取得
  Set chartSheet = Worksheets("グラフ")
  chartSheet.Activate
  Set range = chartSheet.range(Cells(6, 2), Cells(6, 100))
  Set range = range.SpecialCells(xlCellTypeConstants)
  
  ' 削除項目を選択するためのユーザーフォームを作成して表示
  UserForm1.Label1.Caption = "削除したい項目を選択してください"
  For Each cell In range
    UserForm1.ComboBox1.AddItem (cell.Value)
  Next cell
  UserForm1.CommandButton1.Caption = "削除する"
  
  UserForm1.Show
  
End Sub



