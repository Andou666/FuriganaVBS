' 文字列が数値だったらメッセージを表示する
Dim buf
buf = InputBox("入力せよ")

If IsNumeric(buf) Then
    MsgBox "数値に変換可能"
End If