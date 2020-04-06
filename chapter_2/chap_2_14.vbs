' 義務教育期間の判定を追加する
Dim age
age = InputBox("年齢は？")

If IsNumeric(age) Then
    If age < 20 Then
        MsgBox "未成年"
        If age >= 6 And age <= 15 Then
            MsgBox "(義務教育)"
        End If
    ElseIf age < 65 Then
        MsgBox "成人"
    Else
        MsgBox "高齢者"
    End If
End If