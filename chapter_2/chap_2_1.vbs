' isNumeric関数を使ってみよう
Dim buf
buf = InputBox("入力せよ")

' 数値の場合 true
' 数値以外の場合 false
MsgBox isNumeric(buf)