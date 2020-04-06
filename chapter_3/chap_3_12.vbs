' 総当たり戦をつくる
' 同じチーム同士を表示させない
Dim temas, team1, team2
Dim dispMsg
temas = Array("A", "B", "C", "D", "E")
For Each team1 In temas
    For Each team2 In temas
        If team1 <> team2 Then
            dispMsg = dispMsg & team1 & " vs " & team2 & vbCrLf
        End If
    Next
Next

' 一括表示
MsgBox dispMsg
