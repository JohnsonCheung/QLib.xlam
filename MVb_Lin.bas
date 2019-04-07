Attribute VB_Name = "MVb_Lin"
Option Explicit

Function HasDDRmk(A) As Boolean
HasDDRmk = HasSubStr(A, "--")
End Function

Function IsSngTermLin(A) As Boolean
IsSngTermLin = InStr(Trim(A), " ") = 0
End Function

Function IsDDLin(A) As Boolean
IsDDLin = FstTwoChr(LTrim(A)) = "--"
End Function

Function IsDotLin(A) As Boolean
IsDotLin = FstChr(A) = "."
End Function

Function HasLinT1Ay(Lin, T1Ay$()) As Boolean
HasLinT1Ay = HasEle(T1Ay, T1(Lin))
End Function
Function PfxzPfxAy(S, PfxAy)
Dim X
ThwIfNotAy PfxAy, CSub
For Each X In PfxAy
    If HasPfx(S, X) Then PfxzPfxAy = X: Exit Function
Next
End Function

Function PfxzPfxAyPlusSpc(S, PfxAy)
Dim X
ThwIfNotAy PfxAy, CSub
For Each X In PfxAy
    If HasPfx(S, X & " ") Then PfxzPfxAyPlusSpc = X: Exit Function
Next
End Function

Function PfxzPfxApPlusSpc(S, ParamArray PfxAp())
Dim Av(): Av = PfxAp
PfxzPfxApPlusSpc = PfxzPfxAyPlusSpc(S, Av)
End Function

Function PfxzPfxAp(S, ParamArray PfxAp())
Dim Av(): Av = PfxAp
PfxzPfxAp = PfxzPfxAy(S, Av)
End Function

