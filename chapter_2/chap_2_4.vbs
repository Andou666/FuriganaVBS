' Else句を追加してみる
Dim buf
buf = InputBox("入力せよ")

If IsNumeric(buf) Then
    MsgBox buf + 80
Else
    MsgBox "数値計算不可"
End If