VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 削除ボタンが押された時以下の処理を実行
Private Sub CommandButton1_Click()
  ' 削除したい項目名が未選択の時メッセージを表示して項目選択画面に戻る
  If UserForm1.ComboBox1.Value = "" Then
    MsgBox "削除したい項目を選択してください"
    Exit Sub
  End If
  
  ' 選択された項目を本当に削除するか確認するメッセージを表示
  Dim button As Integer
  button = MsgBox("項目名「" & UserForm1.ComboBox1.Value & "」を削除しますか?", vbYesNo + vbQuestion + vbDefaultButton2, "項目削除")
  ' 「いいえ」を選択した時、処理を終了して項目選択画面に戻る
  If button = vbNo Then
    Exit Sub
  
  ' 「はい」が選択された時、以下の処理を実行
  ElseIf button = vbYes Then
    Dim ws As Worksheet
    Dim chartSheet As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Select Case ws.Name
        Case "写真", "コマンドボタン"
          ' 処理対象外なので、処理なし
          
        Case "グラフ"
'        ' グラフのシートに対して以下の処理を実行
'        Dim range As range
'        Dim cell As range

        Set chartSheet = Worksheets("グラフ")

'        ' グラフシート内の異物項目の取得
'        Set range = chartSheet.range(Cells(6, 2), Cells(6, 100))
'        Set range = range.SpecialCells(xlCellTypeConstants)
'
'        For Each cell In range
'          ' グラフシート内の異物項目名とユーザーフォームで選択した異物項目名が一致した時、以下の処理を実行
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
