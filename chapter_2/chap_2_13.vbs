' 「数値判定」のブロック内に「3段階の判定」を書く
Dim age
age = InputBox("年齢は？")

If IsNumeric(age) Then
    If age < 20 Then
        MsgBox "未成年"
    ElseIf age < 65 Then
        MsgBox "成人"
    Else
        MsgBox "高齢者"
    End If
End If