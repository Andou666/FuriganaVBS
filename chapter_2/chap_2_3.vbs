' 数値のときは計算する
Dim buf
buf = InputBox("入力せよ")

If IsNumeric(buf) Then
    MsgBox buf + 80
End If