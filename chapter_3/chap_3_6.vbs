' ���炵���\������
Dim x, y
Dim displayMsg


For x = 1 To 9
    For y = 1 To 9 
        displayMsg = displayMsg & x & "�~" & y & "=" & x * y & " "
    Next
    displayMsg = displayMsg & vbCrLf
Next

' �ꊇ�\��������
MsgBox displayMsg

