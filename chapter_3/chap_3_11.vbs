' ��������������
' �P���ɂ��ׂĂ̑g�ݍ��킹����ׂ�
Dim temas, team1, team2
Dim dispMsg
temas = Array("A", "B", "C", "D", "E")
For Each team1 In temas
    For Each team2 In temas
        dispMsg = dispMsg & team1 & " vs " & team2 & vbCrLf
    Next
Next

' �ꊇ�\��
MsgBox dispMsg
