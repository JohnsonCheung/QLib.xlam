Attribute VB_Name = "MVb_Lin"
Option Explicit

Function HasDDRmk(Lin$) As Boolean
HasDDRmk = HasSubStr(Lin, "--")
End Function

Function IsSngTermLin(Lin$) As Boolean
IsSngTermLin = InStr(Trim(Lin), " ") = 0
End Function

Function IsDDLin(Lin$) As Boolean
IsDDLin = FstTwoChr(LTrim(Lin)) = "--"
End Function

Function IsDotLin(Lin$) As Boolean
IsDotLin = FstChr(Lin) = "."
End Function

Function HitT1Ay(Lin$, T1Sy$()) As Boolean
HitT1Ay = HasEle(T1Sy, T1(Lin))
End Function
Function PfxzPfxAy$(S$, PfxAy$())
Dim Pfx$, I
ThwIfNotAy PfxAy, CSub
For Each I In PfxAy
    Pfx = I
    If HasPfx(S, Pfx) Then PfxzPfxAy = Pfx: Exit Function
Next
End Function

Function PfxzPfxAyPlusSpc(S$, PfxAy$())
Dim X
ThwIfNotAy PfxAy, CSub
For Each X In PfxAy
    If HasPfx(S, X & " ") Then PfxzPfxAyPlusSpc = X: Exit Function
Next
End Function

Function PfxzPfxApPlusSpc(S$, ParamArray PfxAp())
Dim PfxAy$(): PfxAy = SyzAy(PfxAy)
PfxzPfxApPlusSpc = PfxzPfxAyPlusSpc(S, PfxAy)
End Function

Function PfxzPfxAp(S$, ParamArray PfxAp())
Dim PfxAy$(): PfxAy = SyzAy(PfxAy)
PfxzPfxAp = PfxzPfxAy(S, PfxAy)
End Function

