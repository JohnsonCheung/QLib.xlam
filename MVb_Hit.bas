Attribute VB_Name = "MVb_Hit"
Option Explicit
Function HitPfxAy(A, PfxAy) As Boolean
If Si(PfxAy) = 0 Then HitPfxAy = 0: Exit Function
Dim I
For Each I In PfxAy
   If HasPfx(A, I) Then HitPfxAy = True: Exit Function
Next
End Function

Function HitPfxAp(A, ParamArray PfxAp()) As Boolean
Dim Av(): Av = PfxAp
HitPfxAp = HitPfxAy(A, Av)
End Function

Function HitPfxSpc(A, Pfx) As Boolean
HitPfxSpc = HasPfx(A, Pfx & " ")
End Function

Function HitAyElePfx(Ay, ElePfx) As Boolean
Dim S
For Each S In Itr(Ay)
'    If HasPfx(S, Pfx) Then
        HitAyElePfx = True
        Exit Function
'    End If
Next
End Function

Function Has2T(Lin, T1, T2) As Boolean
Dim L$: L = Lin
If ShfT1(L) <> T1 Then Exit Function
If ShfT1(L) <> T2 Then Exit Function
Has2T = True
End Function

Function Has3T(Lin, T1, T2, T3) As Boolean
Dim L$: L = Lin
If ShfT1(L) <> T1 Then Exit Function
If ShfT1(L) <> T2 Then Exit Function
If ShfT1(L) <> T3 Then Exit Function
Has3T = True
End Function

Function Has1T(Lin, T1) As Boolean
Has1T = T1(Lin) = T1
End Function

Function HasT2(Lin, T2) As Boolean
HasT2 = T2(Lin) = T2
End Function
Function HitLikss(S, Likss) As Boolean
HitLikss = HitLikAy(S, SySsl(Likss))
End Function

Function HitLikAy(S, LikeAy$()) As Boolean
Dim Lik
For Each Lik In Itr(LikeAy)
    If S Like Lik Then HitLikAy = True: Exit Function
Next
End Function
Function HitAv(A, Av()) As Boolean
HitAv = HasEle(Av, A)
End Function
Function HitAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
HitAp = HasEle(Av, V)
End Function
Function HitNmStr(V, WhStr$, Optional NmPfx$) As Boolean
HitNmStr = HitNm(V, WhNmzStr(WhStr, NmPfx))
End Function
Function HitNm(V, B As WhNm) As Boolean
HitNm = True
If IsNothing(B) Then Exit Function
If B.IsEmp Then Exit Function
If HitLikAy(V, B.ExlLikAy) Then HitNm = False: Exit Function
If HitRe(V, B.Re) Then Exit Function
If HitLikAy(V, B.LikeAy) Then Exit Function
HitNm = False
End Function
Function HitAy(V, Ay) As Boolean
If Si(Ay) = 0 Then HitAy = True: Exit Function
HitAy = HasEle(Ay, V)
End Function

Private Sub Z_HitPatn()
Dim A$, Patn$
Ept = True: A = "AA": Patn = "AA": GoSub Tst
Ept = True: A = "AA": Patn = "^AA$": GoSub Tst
Exit Sub
Tst:
    Act = HitPatn(A, Patn)
    C
    Return
End Sub

Function HitPatn(A, Patn) As Boolean
Static Re As New RegExp
Re.Pattern = Patn
HitPatn = Re.Test(A)
End Function


Private Sub Z()
Z_HitPatn
MVb_Str_Mch:
End Sub

