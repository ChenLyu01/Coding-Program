Game.autorun
Pathway.clear
Object.clear


a = 0
Do
    Object.create 3, a, a, a
    If a > 6 Then Exit Do
    a = a + 1
Loop Until a < 16