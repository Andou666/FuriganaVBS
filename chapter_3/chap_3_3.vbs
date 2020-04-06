' メッセージの中に回数を入れる
Dim i
Dim text, displayMsg
text = "ハロー！" & vbCrLf

For i = 1 to 10
    displayMsg = displayMsg & i & "回目の" & text
Next

' 10回もメッセージボックスを出したくないので一括で表示させる
' 表示：「n 回目のハロー！」
MsgBox displayMsg

