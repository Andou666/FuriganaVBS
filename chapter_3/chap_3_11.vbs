' ‘“–‚½‚èí‚ğ‚Â‚­‚é
' ’Pƒ‚É‚·‚×‚Ä‚Ì‘g‚İ‡‚í‚¹‚ğ•À‚×‚é
Dim temas, team1, team2
Dim dispMsg
temas = Array("A", "B", "C", "D", "E")
For Each team1 In temas
    For Each team2 In temas
        dispMsg = dispMsg & team1 & " vs " & team2 & vbCrLf
    Next
Next

' ˆêŠ‡•\¦
MsgBox dispMsg
