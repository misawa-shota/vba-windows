Attribute VB_Name = "Module6"
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
