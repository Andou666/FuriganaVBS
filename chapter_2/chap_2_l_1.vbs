' 6歳未満なら「幼児」と表示するマクロ
Dim age
age = InputBox("年齢は？")

If age < 6 Then
    MsgBox "幼児"
End If