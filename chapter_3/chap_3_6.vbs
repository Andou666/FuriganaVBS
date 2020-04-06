' ã„ã„ÇÁÇµÇ≠ï\é¶Ç∑ÇÈ
Dim x, y
Dim displayMsg


For x = 1 To 9
    For y = 1 To 9 
        displayMsg = displayMsg & x & "Å~" & y & "=" & x * y & " "
    Next
    displayMsg = displayMsg & vbCrLf
Next

' àÍäáï\é¶Ç≥ÇπÇÈ
MsgBox displayMsg

