' 2つのIf文を組み合わせる
Dim age
age = InputBox("年齢は？")

If IsNumeric(age) Then
    If age < 20 Then
        MsgBox "未成年"
    End If
End If