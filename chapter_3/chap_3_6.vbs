' 九九らしく表示する
Dim x, y
Dim displayMsg


For x = 1 To 9
    For y = 1 To 9 
        displayMsg = displayMsg & x & "×" & y & "=" & x * y & " "
    Next
    displayMsg = displayMsg & vbCrLf
Next

' 一括表示させる
MsgBox displayMsg

