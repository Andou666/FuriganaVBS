' 曜日を逆順に表示するマクロを書く
Dim wdays, d
Dim dispMsg
wdays = Array("月","火","水","木","金")

For i = 4 To 0 Step -1
    dispMsg = dispMsg & wdays(i) & "曜日" & vbCrLf
Next

' 一括表示
MsgBox dispMsg
