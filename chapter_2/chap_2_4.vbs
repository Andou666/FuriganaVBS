' Else���ǉ����Ă݂�
Dim buf
buf = InputBox("���͂���")

If IsNumeric(buf) Then
    MsgBox buf + 80
Else
    MsgBox "���l�v�Z�s��"
End If