' 残高がゼロになるまで繰り返す
Dim shikin
shikin = 50000
Do While shikin >= 0
    MsgBox "資金：" & shikin
    shikin = shikin - 10800
Loop