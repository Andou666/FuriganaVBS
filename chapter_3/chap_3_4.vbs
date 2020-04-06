' 逆順で繰り返す
Dim i
Dim text, displayMsg
text = "ハロー！" & vbCrLf

For i = 10 to 1 Step -1
    displayMsg = displayMsg & i & "回目の" & text
Next

' 10回もメッセージボックスを出したくないので一括で表示させる
' 表示：「n 回目のハロー！」
MsgBox displayMsg

