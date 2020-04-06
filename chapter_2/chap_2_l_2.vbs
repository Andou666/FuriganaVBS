' 以下のマクロの修正
' Dim age
' age = InputBox("年齢は？")

' If age <= 5 And age >= 65 Then
'     MsgBox "幼児か高齢者"
' End If

Dim age
age = InputBox("年齢は？")

' If age <= 5 And age >= 65 Then
If age <= 5 Or age >= 65 Then
    MsgBox "幼児か高齢者"
End If