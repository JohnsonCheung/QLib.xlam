Attribute VB_Name = "QVb_Lin"
Option Explicit
Private Const CMod$ = "MVb_Lin."
Private Const Asm$ = "QVb"

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
Function PfxzPfxSy$(S$, PfxSy$())
Dim Pfx$, I
ThwIfNotAy PfxSy, CSub
For Each I In PfxSy
    Pfx = I
    If HasPfx(S, Pfx) Then PfxzPfxSy = Pfx: Exit Function
Next
End Function

Function PfxzPfxSyPlusSpc(S$, PfxSy$())
Dim X
ThwIfNotAy PfxSy, CSub
For Each X In PfxSy
    If HasPfx(S, X & " ") Then PfxzPfxSyPlusSpc = X: Exit Function
Next
End Function

Function PfxzPfxApPlusSpc(S$, ParamArray PfxAp())
Dim PfxSy$(): PfxSy = SyzAy(PfxSy)
PfxzPfxApPlusSpc = PfxzPfxSyPlusSpc(S, PfxSy)
End Function

Function PfxzPfxAp(S$, ParamArray PfxAp())
Dim PfxSy$(): PfxSy = SyzAy(PfxSy)
PfxzPfxAp = PfxzPfxSy(S, PfxSy)
End Function

Function IsRmkLin(Lin$) As Boolean
Select Case FstChr(Lin)
Case "#", "@": IsRmkLin = True
End Select
End Function

Function RmvRmkLin(Sy$()) As String()
Dim L, Lin$
For Each L In Itr(Sy)
    Lin = L
    If Not IsRmkLin(Lin) Then
        PushS RmvRmkLin, Lin
    End If
Next
End Function

