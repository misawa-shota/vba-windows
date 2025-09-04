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
    Set chartSheet = Worksheets("グラフ")
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Select Case ws.Name
        Case "写真", "コマンドボタン"
          ' 処理対象外なので、処理なし
          
        Case "グラフ"
        ' グラフシートに以下の処理を実行
        Dim i As Long
        For i = 2 To 100
          ' i + 1番目の値が空の時、i番目の欄を削除して、繰り返し処理を終了する
          If IsEmpty(Cells(6, i + 1).Value) Then
            range(Cells(6, i), Cells(7, i)).Delete
            Exit For
          End If
          ' i番目のセルの値とユーザーフォームで選択された値が一致した時、i番目のセルの値を空にする
          If Cells(6, i).Value = UserForm1.ComboBox1.Value Then
            Cells(6, i).ClearContents
          End If
          ' i番目のセルの値が空の時、i + 1番目の値をi番目に移動して、i + 1番目の値を空にする
          If IsEmpty(Cells(6, i).Value) Then
             Cells(6, i + 1).Offset(0, -1).Value = Cells(6, i + 1).Value
             Cells(6, i + 1).ClearContents
          End If
        Next i
        
        Case Else
        ' 商品のシートに以下の処理を実行
        Dim n As Long
        For n = 4 To 100
          ' ユーザーフォームで選択された異物項目名に一致する項目を削除して、それ以降の欄を左に移動させる
          If Cells(6, n).Value = UserForm1.ComboBox1.Value Then
            range(Cells(6, n), Cells(35, n)).Delete Shift:=xlToLeft
            ' 削除する異物項目が最後の項目の時、その欄の左の枠線を太くする処理
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
