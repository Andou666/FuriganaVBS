' 同じメッセージを10回表示する
Dim i
Dim text, displayMsg
text = "ハロー！" & vbCrLf

For i = 1 to 10
    displayMsg = displayMsg & text
Next

' 10回もメッセージボックスを出したくないので一括で表示させる
MsgBox displayMsg

