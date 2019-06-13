Attribute VB_Name = "QVb_Lin_Shf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Lin_Shf."
Private Const Asm$ = "QVb"
Function ShfDotSeg$(OLin$)
ShfDotSeg = ShfBef(OLin, ".")
End Function
Function ShfBktStr$(OLin$)
ShfBktStr = BetBkt(OLin)
OLin = AftBkt(OLin$)
End Function
Function RmvChr$(S, ChrLis$) ' Rmv fst chr if it is in ChrLis
If HasSubStr(ChrLis, FstChr(S)) Then
    RmvChr = RmvFstChr(S)
Else
    RmvChr = S
End If
End Function
Function TakChr$(S, ChrLis$) ' Ret fst chr if it is in ChrLis
If HasSubStr(ChrLis, FstChr(S)) Then TakChr = FstChr(S)
End Function

Function RmvChrzSfx$(S, ChrLis$)
If HasSubStr(ChrLis, LasChr(S)) Then
    RmvChrzSfx = RmvLasChr(S)
Else
    RmvChrzSfx = S
End If
End Function
Function ShfChr$(OLin$, ChrList$)
Dim C$: C = TakChr(OLin, ChrList)
If C = "" Then Exit Function
ShfChr = C
OLin = Mid(OLin, 2)
End Function
Function ShfTermX(OLin$, TermX$) As Boolean
If T1(OLin) <> TermX Then Exit Function
ShfTermX = True
OLin = RmvT1(OLin)
End Function
Function ShfAs(OLin$) As Boolean
ShfAs = ShfTermX(OLin, "As")
End Function
Function ShfEq(OLin$) As Boolean
ShfEq = ShfTermX(OLin, "=")
End Function
Function ShfTy(OLin$) As Boolean
ShfTy = ShfTermX(OLin, "Ty")
End Function
Function ShfBkt(OLin$) As Boolean
ShfBkt = ShfPfx(OLin, "()")
End Function
Function ShfPfx(OLin$, Pfx$) As Boolean
If HasPfx(OLin, Pfx) Then
    OLin = RmvPfx(OLin, Pfx)
    ShfPfx = True
End If
End Function

Function ShfPfxSpc(OLin$, Pfx$) As Boolean
If HitPfxSpc(OLin, Pfx) Then
    OLin = Mid(OLin, Len(Pfx) + 2)
    ShfPfxSpc = True
End If
End Function

Private Sub Z_ShfBktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Private Property Get Z_ShfPfx()
Dim O$: O = "AA{|}BB "
Ass ShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Property



Private Sub ZZ()
Z_ShfBktStr

MVb_Lin_Shf:
End Sub
