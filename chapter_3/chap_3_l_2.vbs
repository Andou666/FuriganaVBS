' �j�����t���ɕ\������}�N��������
Dim wdays, d
Dim dispMsg
wdays = Array("��","��","��","��","��")

For i = 4 To 0 Step -1
    dispMsg = dispMsg & wdays(i) & "�j��" & vbCrLf
Next

' �ꊇ�\��
MsgBox dispMsg
