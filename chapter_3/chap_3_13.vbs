' ‘“–‚½‚èí‚ğ‚Â‚­‚é
' “¯‚¶‘Îí‘g‚İ‡‚í‚¹‚ğÈ‚­
Dim teams, team1, team2Top, i
Dim dispMsg
teams = Array("A", "B", "C", "D", "E")
team2Top = 1
For Each team1 In teams
    For i = team2Top To 4
        dispMsg = dispMsg & team1 & " vs " & teams(i) & vbCrLf
    Next
    team2Top = team2Top + 1
Next

' ˆêŠ‡•\¦
MsgBox dispMsg
