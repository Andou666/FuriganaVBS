' 東西南北を表示するマクロを書く
Dim disrecs, d
Dim dispMsg
disrecs = Array("東","西","南","北")

For Each d In disrecs
    dispMsg = dispMsg & d & vbCrLf
Next

' 一括表示
MsgBox dispMsg
