' 九九の計算をしてみる
Dim x, y
Dim text, displayMsg


For x = 1 To 9
    For y = 1 To 9 
        displayMsg = displayMsg & x * y & " "
    Next
    displayMsg = displayMsg & vbCrLf
Next

' 一括表示させる
MsgBox displayMsg

