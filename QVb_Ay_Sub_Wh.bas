Attribute VB_Name = "QVb_Ay_Sub_Wh"
Option Explicit
Private Const CMod$ = "MVb_Ay_Sub_Wh."
Private Const Asm$ = "QVb"

Sub AssDup(Ay, Fun$)
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = AywDup(Ay)
If Si(Dup) = 0 Then Exit Sub
Thw Fun, "There are dup in array", "Dup Ay", Dup, Ay
End Sub
Function AywIxCnt(Ay, Ix, Cnt)
Dim J&
Dim O: O = Ay: Erase O
For J = 0 To Cnt - 1
    Push O, Ay(Ix + J)
Next
AywIxCnt = O
End Function

Function AywBetEle(Ay, FmEle, ToEle)
Dim O: O = Resi(Ay)
Dim Inside As Boolean, I
For Each I In Itr(Ay)
    If Inside Then
        If I = ToEle Then
            AywBetEle = O
            Exit Function
        End If
        PushI O, I
    Else
        If I = FmEle Then
            Inside = True
        End If
    End If
Next
If Inside Then
    Thw CSub, "FmEle is found, but no ToEle", "Ay FmEle ToEle", Ay, FmEle, ToEle
End If
End Function
Function AywDist(Ay, Optional IgnCas As Boolean)
AywDist = IntozItr(Resi(Ay), CntDic(Ay).Keys)
End Function
Private Sub Z_FmtCntDic()
Dim Ay
GoSub ZZ
Exit Sub
ZZ:
    Ay = Array(1, 2, 2, 2, 3, "skldflskdfsdklf" & vbCrLf & "ksdlfj")
    Brw FmtCntDic(Ay)
    
End Sub


Function AywDistT1(Sy$()) As String()
AywDistT1 = AywDist(T1Ay(Sy))
End Function

Function AywDup(Ay, Optional C As VbCompareMethod = vbTextCompare)
Dim O: O = Resi(Ay)
Dim D As Dictionary: Set D = CntDic(Ay, EiCntDup, C)
AywDup = IntozItr(O, D.Keys)
End Function

Function AywFmIx(Ay, FmIx)
Dim O: O = Ay: Erase O
If 0 <= FmIx And FmIx <= UB(Ay) Then
    Dim J&
    For J = FmIx To UB(Ay)
        Push O, Ay(J)
    Next
End If
AywFmIx = O
End Function

Function AywFE(Ay, FmIx, EIx)
AywFE = AywFT(Ay, FmIx, EIx - 1)
End Function

Function AywFT(Ay, FmIx, ToIx)
Dim J&
AywFT = Resi(Ay)
For J = FmIx To ToIx
    Push AywFT, Ay(J)
Next
End Function
Function AywFstUEle(Ay, U)
Dim O: O = Ay
ReDim Preserve O(U)
AywFstUEle = O
End Function

Function AywFstNEle(Ay, N)
AywFstNEle = AywFstUEle(Ay, N - 1)
End Function

Function AywFEIx(Ay, B As FEIx)
AywFEIx = AywFE(Ay, B.FmIx, B.EIx)
End Function

Function IsOutRange(Ixy, U&) As Boolean
Dim Ix
For Each Ix In Itr(Ixy)
    If 0 > Ix Then IsOutRange = True: Exit Function
    If Ix > U Then IsOutRange = True: Exit Function
Next
End Function
Function AywIxyzMust(Ay, Ixy&())
If IsOutRange(Ixy, UB(Ay)) Then Thw CSub, "Some element in Ixy is outsize Ay", "UB(Ay) Ixy", UB(Ay), Ixy
Dim U&: U = UB(Ay)
Dim O: O = Resi(Ay)
Dim Ix
For Each Ix In Itr(Ixy)
    Push O, Ay(Ix)
Next
AywIxyzMust = O
End Function
Function AywInAset(Ay, B As Aset)
AywInAset = Resi(Ay)
Dim I
For Each I In Itr(Ay)
    If Ay.Has(I) Then Push AywInAset, I
Next
End Function
Function AywIxy(Ay, Ixy&())
Dim U&: U = UB(Ixy)
Dim O: O = Resi(Ay)
ReDim Preserve O(U)
Dim Ix, J&
For Each Ix In Itr(Ixy)
    If 0 > Ix Or Ix > U Then
        Push O(J), Ay(Ix)
    End If
    J = J + 1
Next
AywIxy = O
End Function

Function SywLik(Sy$(), Lik$) As String()
Dim I
For Each I In Itr(Sy)
    If I Like Lik Then PushI SywLik, I
Next
End Function

Function SywLikSy(Ay, LikeAy$()) As String()
Dim I, Lik
For Each I In Itr(Ay)
    For Each Lik In LikeAy
        If I Like Lik Then
            PushI SywLikSy, I
            Exit For
        End If
    Next
Next
End Function
Function IsEmpWhNm(Ay As WhNm) As Boolean
IsEmpWhNm = True
If IsNothing(Ay) Then Exit Function
With Ay
    If IsNothing(.Re) Then
        If Si(.ExlLikSy) = 0 Then
            If Si(.LikeAy) = 0 Then
                Exit Function
            End If
        End If
    End If
End With
IsEmpWhNm = True
End Function
Function AywNmStr(Sy$(), WhStr$, Optional NmPfx$) As String()
AywNmStr = AywNm(Sy, WhNmzStr(WhStr, NmPfx))
End Function

Function AywNm(Sy$(), B As WhNm) As String()
Dim I
For Each I In Itr(Sy)
    If HitNm(CStr(I), B) Then PushI AywNm, I
Next
End Function

Function AyePfx(Sy$(), Pfx$) As String()
Dim I
For Each I In Itr(Sy)
    If Not HasPfx(CStr(I), Pfx) Then PushI AyePfx, I
Next
End Function

Function AywPred(Ay, B As IPred)
Dim I, O, X
AywPred = Resi(Ay)
For Each I In Itr(Ay)
    If B.Pred(I) Then
        Push AywPred, I
    End If
Next
End Function
Function PatnPred(Patn$) As IPred
Dim O As New PredzPatn
O.Init Patn
Set PatnPred = O
End Function
Function SywPatn(Sy$(), Patn$) As String()
If Si(Sy) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then SywPatn = Sy: Exit Function
SywPatn = SywPred(Sy, PatnPred(Patn))
End Function
Function SyePred(Sy$(), P As IPred) As String()
Dim I
For Each I In Itr(Sy)
    If Not P.Pred(I) Then
        PushI SyePred, I
    End If
Next
End Function

Function SywPred(Sy$(), P As IPred) As String()
Dim I
For Each I In Itr(Sy)
    If P.Pred(I) Then
        PushI SywPred, I
    End If
Next
End Function
Function SywPatnExl(Sy$(), Patn$, ExlLikss$) As String()
SywPatnExl = SyeLikss(SywPatn(Sy, Patn), ExlLikss)
End Function
Function SyeLikss(Sy$(), ExlLikss$) As String()
SyeLikss = SyePred(Sy, PredzLikss(ExlLikss))
End Function
Function PredzLikss(Likss$) As IPred


End Function
Function IxyzAyPatn(Ay, Patn$) As Long()
IxyzAyPatn = IxyzAyRe(Ay, RegExp(Patn))
End Function
Function IxyzAyRe(Ay, B As RegExp) As Long()
If Si(Ay) = 0 Then Exit Function
Dim I, O&(), J&
For Each I In Ay
    If B.Test(I) Then Push O, J
    J = J + 1
Next
IxyzAyRe = O
End Function

Function SywPfx(Sy$(), Pfx$) As String()
Dim I
For Each I In Itr(Sy)
    If HasPfx(CStr(I), Pfx) Then PushI SywPfx, I
Next
End Function

Function AywPredFalse(Ay, P As IPred)
Dim X
AywPredFalse = Resi(Ay)
For Each X In Itr(Ay)
    If Not P.Pred(X) Then
        Push AywPredFalse, X
    End If
Next
End Function

Function AywPredXAB(Ay, P As IPredXAB, A, B)
Dim X
AywPredXAB = Resi(Ay)
For Each X In Itr(Ay)
    If P.PredXAB(X, A, B) Then
        Push AywPredXAB, X
    End If
Next
End Function


Function AywPredXAP(Ay, PredXAP$, ParamArray Ap())
AywPredXAP = Resi(Ay)
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In Itr(Ay)
    Asg I, Av(0)
    If RunAv(PredXAP, Av) Then
        Push AywPredXAP, I
    End If
Next
End Function

Function AywPredXP(Ay, XP$, P)
Dim X
AywPredXP = Resi(Ay)
For Each X In Itr(Ay)
    If Run(XP, X, P) Then
        Push AywPredXP, X
    End If
Next
End Function

Function AywPredXPNot(Ay, XP$, P)
Dim X
AywPredXPNot = Resi(Ay)
For Each X In Itr(Ay)
    If Not Run(XP, X, P) Then
        Push AywPredXPNot, X
    End If
Next
End Function

Function AywRe(Ay, Re As RegExp) As String()
If IsNothing(Re) Then AywRe = SyzAy(Ay): Exit Function
Dim X
For Each X In Itr(Ay)
    If Re.Test(X) Then PushI AywRe, X
Next
End Function
Function AywRmvEle(Ay, Ele)
AywRmvEle = Resi(Ay)
Dim I
For Each I In Itr(Ay)
    If I <> Ele Then PushI AywRmvEle, I
Next
End Function
Function ItrzSywRmvT1(Sy$(), T1)
Asg Itr(SywRmvT1(Sy, T1)), ItrzSywRmvT1
End Function

Function ItrzSsLin(Ssl$)
Asg SyzSS(Ssl), ItrzSsLin
End Function

Function SywRmvT1(Sy$(), T1) As String()
SywRmvT1 = RmvT1zSy(AywT1(Sy, T1))
End Function

Function SywRmvTT(Sy$(), T1, T2) As String()
SywRmvTT = RmvTTzSy(SywTT(Sy, T1, T2))
End Function

Function AySkip(Ay, Optional SkipN& = 1)
Dim O: O = Resi(Ay)
Dim J&
For J = SkipN To UB(Ay)
    Push O, Ay(J)
Next
End Function

Function AywSfx(Sy$(), Sfx$) As String()
Dim I
For Each I In Itr(Sy)
    If HasSfx(CStr(I), Sfx) Then PushI AywSfx, I
Next
End Function

Function AywSingleEle(Ay)
Dim O: O = Ay: Erase O
Dim CntDry(): CntDry = CntDryzAy(Ay)
If Si(CntDry) = 0 Then
    AywSingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDry
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
AywSingleEle = O
End Function

Function AywSng(Ay)
AywSng = MinusAy(Ay, AywDup(Ay))
End Function

Function AywSngEle(Ay)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = Resi(Ay)
Dim K, D As Dictionary
Set D = CntDic(Ay)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AywT1(Ay, T1) As String()
Dim L$, I
For Each I In Itr(Sy)
    L = I
    If HasT1(L, T1) Then
        PushI AywT1, L
    End If
Next
End Function

Function AywT1InAy(Ay, InAy) As String()
If Si(Sy) = 0 Then Exit Function
Dim O$(), L
For Each L In Sy
    If HasEle(InAy, T1(CStr(L))) Then Push O, L
Next
AywT1InAy = O
End Function

Function AywT1SelRst(Sy$(), T1) As String()
Dim I, L$
For Each I In Itr(Sy)
    L = I
    If ShfT1(L) = T1 Then PushI AywT1SelRst, L
Next
End Function

Function SywTT(Sy$(), T1, T2) As String()
Dim I, L$
For Each I In Itr(Sy)
    L = I
    If HasTT(L, T1, T2) Then PushI SywTT, L
Next
End Function

Function SywTTSelRst(Sy$(), T1, T2) As String()
Dim L$, I, X1$, X2$, Rst$
For Each I In Itr(Sy)
    L = I
    AsgN2tRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI SywTTSelRst, Rst
        End If
    End If
Next
End Function

Private Sub ZZ()
Dim Ay As Variant
Dim B$
Dim C As Boolean
Dim D&
Dim E As FEIx
Dim F$()
Dim G As WhNm
Dim H()
Dim I As RegExp
AywDist Ay, C
FmtCntDic Ay
End Sub
