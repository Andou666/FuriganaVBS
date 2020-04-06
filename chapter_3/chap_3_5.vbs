' ‹ã‹ã‚ÌŒvŽZ‚ð‚µ‚Ä‚Ý‚é
Dim x, y
Dim text, displayMsg


For x = 1 To 9
    For y = 1 To 9 
        displayMsg = displayMsg & x * y & " "
    Next
    displayMsg = displayMsg & vbCrLf
Next

' ˆêŠ‡•\Ž¦‚³‚¹‚é
MsgBox displayMsg

