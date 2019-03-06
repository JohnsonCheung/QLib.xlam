Attribute VB_Name = "MVb_Lin_Shf"
Option Explicit

Function ShfBktStr$(OLin$)
Dim O$
O = TakBetBkt(OLin): If O = "" Then Exit Function
ShfBktStr = O
OLin = TakAftBkt(OLin)
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
Function ShfChr$(OLin, ChrList$)
Dim C$: C = TakChr(OLin, ChrList)
If C = "" Then Exit Function
ShfChr = C
OLin = Mid(OLin, 2)
End Function

Function ShfPfx(OLin, Pfx) As Boolean
If HasPfx(OLin, Pfx) Then
    OLin = RmvPfx(OLin, Pfx)
    ShfPfx = True
End If
End Function

Function ShfPfxSpc(OLin, Pfx) As Boolean
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



Private Sub Z()
Z_ShfBktStr

MVb_Lin_Shf:
End Sub