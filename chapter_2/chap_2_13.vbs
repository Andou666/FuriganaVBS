' �u���l����v�̃u���b�N���Ɂu3�i�K�̔���v������
Dim age
age = InputBox("�N��́H")

If IsNumeric(age) Then
    If age < 20 Then
        MsgBox "�����N"
    ElseIf age < 65 Then
        MsgBox "���l"
    Else
        MsgBox "�����"
    End If
End If