Attribute VB_Name = "QVb_Ay_Sub_Wh"
Option Compare Text
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
Dim O: O = AyzReSi(Ay)
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
AywDist = IntozItr(AyzReSi(Ay), CntDic(Ay).Keys)
End Function
Private Sub Z_FmtCntDic()
Dim Ay
GoSub ZZ
Exit Sub
ZZ:
    Ay = Array(1, 2, 2, 2, 3, "skldflskdfsdklf" & vbCrLf & "ksdlfj")
    Brw FmtCntDic(Ay)
    
End Sub


Function AywDistT1(Ay) As String()
AywDistT1 = AywDist(T1Ay(Ay))
End Function

Function AywDup(Ay, Optional C As VbCompareMethod = vbTextCompare)
Dim O: O = AyzReSi(Ay)
Dim D As Dictionary: Set D = CntDic(Ay, EiCntDup, C)
AywDup = IntozItr(O, D.Keys)
End Function
Function AywNonEmp(Ay)
AywNonEmp = Ay: Erase AywNonEmp
Dim I
For Each I In Ay
    If Not IsEmpty(I) Then
        PushI AywNonEmp, I
    End If
Next
End Function
Function AywFmIx(Ay, FmIx)
AywFmIx = Ay: Erase AywFmIx
If 0 <= FmIx And FmIx <= UB(Ay) Then
    Dim J&
    For J = FmIx To UB(Ay)
        Push AywFmIx, Ay(J)
    Next
End If
End Function

Function AywFE(Ay, FmIx, EIx)
AywFE = AywFT(Ay, FmIx, EIx - 1)
End Function

Function AywFT(Ay, FmIx, ToIx)
Dim J&
AywFT = AyzReSi(Ay)
For J = FmIx To ToIx
    Push AywFT, Ay(J)
Next
End Function
Function AywFstUEle(Ay, U)
If U > UB(Ay) Then AywFstUEle = Ay: Exit Function
Dim O: O = Ay
ReDim Preserve O(U)
AywFstUEle = O
End Function

Function FstNEle(Ay, N)
FstNEle = AywFstUEle(Ay, N - 1)
End Function

Function AywFei(Ay, B As Fei)
AywFei = AywFE(Ay, B.FmIx, B.EIx)
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
Dim O: O = AyzReSi(Ay)
Dim Ix
For Each Ix In Itr(Ixy)
    Push O, Ay(Ix)
Next
AywIxyzMust = O
End Function

Function AywInAset(Ay, B As Aset)
AywInAset = AyzReSi(Ay)
Dim I
For Each I In Itr(Ay)
    If Ay.Has(I) Then Push AywInAset, I
Next
End Function

Function AywIxy(Ay, Ixy&())
Dim U&: U = UB(Ixy)
Dim O: O = AyzReSi(Ay)
ReDim Preserve O(U)
Dim Ix, J&
For Each Ix In Itr(Ixy)
    If IsObject(Ay(Ix)) Then
        Set O(J) = Ay(Ix)
    Else
        O(J) = Ay(Ix)
    End If
    J = J + 1
Next
AywIxy = O
End Function

Function AywIxyAlwE(Ay, Ixy&())
Dim U&: U = UB(Ixy)
Dim O: O = AyzReSi(Ay)
ReDim Preserve O(U)
Dim Ix, J&
For Each Ix In Itr(Ixy)
    If Ix >= 0 Then
        If IsObject(Ay(Ix)) Then
            Set O(J) = Ay(Ix)
        Else
            O(J) = Ay(Ix)
        End If
    End If
    J = J + 1
Next
AywIxyAlwE = O
End Function

Function AywLik(Ay, Lik) As String()
Dim I
For Each I In Itr(Ay)
    If I Like Lik Then PushI AywLik, I
Next
End Function
Function IsLikAy(S, LikAy$()) As Boolean
Dim Lik
For Each Lik In LikAy
    If S Like Lik Then IsLikAy = True: Exit Function
Next
End Function
Function AywLikAy(Ay, LikAy$()) As String()
Dim I, Lik
For Each I In Itr(Ay)
    If IsLikAy(I, LikAy) Then PushI AywLikAy, I
Next
End Function

Function AywNmStr(Ay, WhNmStr$) As String()
AywNmStr = AywNm(Ay, WhNmzS(WhNmStr))
End Function

Function AywIsNm(Ay) As String()
AywIsNm = AywPred(Ay, PredIsNm)
End Function

Function AywNm(Ay, B As WhNm) As String()
Dim I
For Each I In Itr(Ay)
    If HitNm(I, B) Then PushI AywNm, I
Next
End Function

Function AyePfx(Ay, Pfx$) As String()
Dim I
For Each I In Itr(Ay)
    If Not HasPfx(I, Pfx) Then PushI AyePfx, I
Next
End Function

Function AywPred(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If P.Pred(I) Then
        PushI AywPred, I
    End If
Next
End Function

Function PatnPred(Patn$) As IPred
Dim O As New PredPatn
O.Init Patn
Set PatnPred = O
End Function

Function AywPatn1(Ay, Patn$) As Variant()
If Si(Ay) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AywPatn1 = Ay: Exit Function
Dim Re As RegExp: Set Re = RegExp(Patn)
Dim I: For Each I In Itr(Ay)
    If IsStr(I) Then
        If Re.Test(I) Then PushI AywPatn1, I
    End If
Next
End Function

Function AywPatn(Ay, Patn$) As String()
If Si(Ay) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AywPatn = Ay: Exit Function
AywPatn = AywPred(Ay, PatnPred(Patn))
End Function
Function AywPatnAy(Ay, PatnAy$()) As Variant()
If Si(Ay) = 0 Then Exit Function
Stop
End Function
Function AyePred(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If Not P.Pred(I) Then
        PushI AyePred, I
    End If
Next
End Function

Function AywPatnExl(Ay, Patn$, ExlLikss$) As String()
AywPatnExl = AyeLikss(AywPatn(Ay, Patn), ExlLikss)
End Function
Function AyeLikss(Ay, ExlLikss$) As String()
Stop
'AyeLikss = AyePred(Ay, PredzIsLikss(ExlLikss))
End Function
Function PredzLikss(Likss$) As IPred
Dim O As New PredLikAy
O.Init SyzSS(Likss)
Set PredzLikss = O
End Function
Function IxyzSubAy(Ay, SubAy, Optional ThwNFnd As Boolean) As Long()
Dim E, Ix&
For Each E In SubAy
    Ix = IxzAy(Ay, E)
    If ThwNFnd Then
        If Ix = -1 Then
            Thw CSub, "Ele in SubAy not found in Ay", "Ele SubAy Ay", E, SubAy, Ay
        End If
    End If
    PushI IxyzSubAy, Ix
Next
End Function

Function IxyzAyPatn(Ay, Patn$) As Long()
IxyzAyPatn = IxyzAyRe(Ay, RegExp(Patn))
End Function
Function IxyzCC(D As Drs, CC$) As Long()
IxyzCC = IxyzAyCC(D.Fny, CC)
End Function

Function IxyzAyCC(Ay, CC$) As Long()
Dim C
For Each C In Itr(SyzSS(CC))
    PushI IxyzAyCC, IxzAy(Ay, C)
Next
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

Function AywPredFalse(Ay, P As IPred)
Dim X
AywPredFalse = AyzReSi(Ay)
For Each X In Itr(Ay)
    If Not P.Pred(X) Then
        Push AywPredFalse, X
    End If
Next
End Function

Function AywPredXAP(Ay, PredXAP$, ParamArray Ap())
AywPredXAP = AyzReSi(Ay)
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

Function AywPredXP(Ay, Xp$, P)
Dim X
AywPredXP = AyzReSi(Ay)
For Each X In Itr(Ay)
    If Run(Xp, X, P) Then
        Push AywPredXP, X
    End If
Next
End Function

Function AywPredXPNot(Ay, Xp$, P)
Dim X
AywPredXPNot = AyzReSi(Ay)
For Each X In Itr(Ay)
    If Not Run(Xp, X, P) Then
        Push AywPredXPNot, X
    End If
Next
End Function

Function AywRe(Ay, Re As RegExp) As String()
Dim S
For Each S In Itr(Ay)
    If Re.Test(S) Then PushI AywRe, S
Next
End Function
Function AywRmvEle(Ay, ele)
AywRmvEle = AyzReSi(Ay)
Dim I
For Each I In Itr(Ay)
    If I <> ele Then PushI AywRmvEle, I
Next
End Function
Function ItrzAywRmvT1(Ay, T1)
Asg Itr(AywRmvT1(Ay, T1)), ItrzAywRmvT1
End Function

Function ItrzSS(SS)
Asg Itr(SyzSS(SS)), ItrzSS
End Function

Function AywRmvT1(Ay, T1) As String()
AywRmvT1 = RmvT1zAy(AywT1(Ay, T1))
End Function

Function AywRmvTT(Ay, T1, T2) As String()
AywRmvTT = RmvTTzAy(AywTT(Ay, T1, T2))
End Function

Function AySkip(Ay, Optional SkipN& = 1)
Dim O: O = AyzReSi(Ay)
Dim J&
For J = SkipN To UB(Ay)
    Push O, Ay(J)
Next
End Function

Function AywSfx(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    If HasSfx(I, Sfx) Then PushI AywSfx, I
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
Dim O: O = AyzReSi(Ay)
Dim K, D As Dictionary
Set D = CntDic(Ay)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AywT1(Ay, T1) As String()
Dim L$, I
For Each I In Itr(Ay)
    L = I
    If HasT1(L, T1) Then
        PushI AywT1, L
    End If
Next
End Function

Function AywT1InAy(Ay, InAy) As String()
If Si(Ay) = 0 Then Exit Function
Dim O$(), L
For Each L In Ay
    If HasEle(InAy, T1(CStr(L))) Then Push O, L
Next
AywT1InAy = O
End Function

Function AywT1SelRst(Ay, T1) As String()
Dim I, L$
For Each I In Itr(Ay)
    L = I
    If ShfT1(L) = T1 Then PushI AywT1SelRst, L
Next
End Function

Function AywTT(Ay, T1, T2) As String()
Dim I, L$
For Each I In Itr(Ay)
    L = I
    If HasTT(L, T1, T2) Then PushI AywTT, L
Next
End Function

Function AywTTSelRst(Ay, T1, T2) As String()
Dim L$, I, X1$, X2$, Rst$
For Each I In Itr(Ay)
    L = I
    AsgN2tRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI AywTTSelRst, Rst
        End If
    End If
Next
End Function

Private Sub ZZ()
Dim Ay As Variant
Dim B$
Dim C As Boolean
Dim D&
Dim E As Fei
Dim F$()
Dim G As WhNm
Dim H()
Dim I As RegExp
AywDist Ay, C
FmtCntDic Ay
End Sub

