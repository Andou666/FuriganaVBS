' 総当たり戦をつくる
' 単純にすべての組み合わせを並べる
Dim temas, team1, team2
Dim dispMsg
temas = Array("A", "B", "C", "D", "E")
For Each team1 In temas
    For Each team2 In temas
        dispMsg = dispMsg & team1 & " vs " & team2 & vbCrLf
    Next
Next

' 一括表示
MsgBox dispMsg
