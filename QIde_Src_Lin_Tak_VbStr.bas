Attribute VB_Name = "QIde_Src_Lin_Tak_VbStr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Src_Lin_Tak_VbStr."
Private Const Asm$ = "QIde"
Function TakVbStr$(S)
If FstChr(S) <> """" Then Exit Function
Dim P%: P = EndPos(2, S, 0)
If P = 0 Then Stop: Exit Function
TakVbStr = Replace(Mid(S, 2, P - 2), """""", """")
End Function
Private Function EndPos%(Fm%, S, Lvl%)
If Lvl > 1000 Then ThwLoopingTooMuch CSub
Dim P%: P = InStr(Fm, S, """"): If P = 0 Then Exit Function
If Mid(S, P + 1, 1) <> """" Then EndPos = P: Exit Function
EndPos = EndPos(P + 2, S, Lvl + 1)
End Function
Private Sub Z_TakVbStr()
Dim S$
'GoSub T1
GoSub T2
'GoSub T3
Exit Sub
T1: S = """aa""": Ept = "aa":       GoTo Tst
T2: S = """aa""""""": Ept = "aa""": GoTo Tst
T3: S = """aa""": Ept = "aa":       GoTo Tst
Tst: Act = TakVbStr(S): Debug.Assert Act = Ept: Return
End Sub
