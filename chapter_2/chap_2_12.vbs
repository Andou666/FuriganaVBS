' Not演算子を使って False のときだけ実行する
Dim age
age = InputBox("年齢は？")

If Not IsNumeric(age) Then
    MsgBox "数値を入力して"
End If