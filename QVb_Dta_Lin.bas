Attribute VB_Name = "QVb_Dta_Lin"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Lin."
Private Const Asm$ = "QVb"

Function HasDDRmk(Lin) As Boolean
HasDDRmk = HasSubStr(Lin, "--")
End Function

Function HitT1Ay(Lin, T1Ay$()) As Boolean
HitT1Ay = HasEle(T1Ay, T1(Lin))
End Function
Function PfxzPfxSy$(S, PfxSy$())
Dim Pfx$, I
ThwIf_NotAy PfxSy, CSub
For Each I In PfxSy
    Pfx = I
    If HasPfx(S, Pfx) Then PfxzPfxSy = Pfx: Exit Function
Next
End Function

Function IsLinVbRmk(L) As Boolean
IsLinVbRmk = FstChr(LTrim(L)) = "'"
End Function

Function RmvRmk$(Lin)
RmvRmk = BefOrAll(Lin, "--", True)
End Function
Function RmvRmkzLy(Ly$()) As String()
Dim L
For Each L In Itr(Ly)
    PushI RmvRmkzLy, RmvRmk(L)
Next
End Function


'
