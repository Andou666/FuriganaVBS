' ElseIf句の書き方
Dim age
age = InputBox("年齢は？")

If age < 20 Then
    MsgBox "未成年"
ElseIf age < 65 Then
    MsgBox "成人"
Else
    MsgBox "高齢者"
End If