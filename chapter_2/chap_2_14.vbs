' �`��������Ԃ̔����ǉ�����
Dim age
age = InputBox("�N��́H")

If IsNumeric(age) Then
    If age < 20 Then
        MsgBox "�����N"
        If age >= 6 And age <= 15 Then
            MsgBox "(�`������)"
        End If
    ElseIf age < 65 Then
        MsgBox "���l"
    Else
        MsgBox "�����"
    End If
End If