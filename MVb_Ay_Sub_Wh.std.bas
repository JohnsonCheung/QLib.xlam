Attribute VB_Name = "MVb_Ay_Sub_Wh"
Option Explicit

Sub AyDupAss(A, Fun$, Optional IgnCas As Boolean)
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = AywDup(A, IgnCas)
If Sz(Dup) = 0 Then Exit Sub
Thw Fun, "There are dup in array", "Dup Ay", Dup, A
End Sub

Function AywDist(A, Optional IgnCas As Boolean)
AywDist = IntozAy(AyCln(A), CntDic(A, IgnCas).Keys)
End Function

Function AywDistFmt(A) As String()
Dim D As Dictionary
Set D = CntDic(A)
AywDistFmt = FmtDic(D)
End Function

Function AywDistSy(A) As String()
AywDistSy = CvSy(AywDist(A))
End Function

Function AywDistT1(A) As String()
AywDistT1 = AywDist(AyTakT1(A))
End Function

Function AywDup(A, Optional IgnCas As Boolean)
Dim D As Dictionary, I
AywDup = AyCln(A)
Set D = CntDic(A, IgnCas)
For Each I In Itr(A)
    If D(I) > 1 Then
        Push AywDup, I
    End If
Next
End Function

Function AywFmIx(A, FmIx)
Dim O: O = A: Erase O
If 0 <= FmIx And FmIx <= UB(A) Then
    Dim J&
    For J = FmIx To UB(A)
        Push O, A(J)
    Next
End If
AywFmIx = O
End Function

Function AywFT(A, FmIx, ToIx)
Dim J&, O
O = AyCln(A)
For J = FmIx To ToIx
    Push O, A(J)
Next
AywFT = O
End Function
Function AywFstUEle(Ay, U)
Dim O: O = Ay
ReDim Preserve O(U)
AywFstUEle = O
End Function

Function AywFstNEle(Ay, N)
Dim O: O = Ay
ReDim Preserve O(N - 1)
AywFstNEle = O
End Function

Function AywFTIx(A, B As FTIx)
AywFTIx = AywFT(A, B.FmIx, B.ToIx)
End Function
Function IsOutRange(IxAy, U&) As Boolean
Dim Ix
For Each Ix In Itr(IxAy)
    If 0 > Ix Then IsOutRange = True: Exit Function
    If Ix > U Then IsOutRange = True: Exit Function
Next
End Function
Function AywIxAyzMust(Ay, IxAy)
If IsOutRange(IxAy, UB(Ay)) Then Thw CSub, "Some element in IxAy is outsize Ay", "UB(Ay) IxAy", UB(Ay), IxAy
Dim U&: U = UB(Ay)
Dim O: O = AyCln(Ay)
Dim Ix
For Each Ix In Itr(IxAy)
    Push O, Ay(Ix)
Next
AywIxAyzMust = O
End Function
Function AywIxAy(A, IxAy)
Dim U&: U = UB(A)
Dim O: O = AyCln(A)
Dim Ix
For Each Ix In Itr(IxAy)
    If 0 > Ix Or Ix > U Then
        ReDim Preserve O(Sz(O))
    Else
        Push O, A(Ix)
    End If
Next
AywIxAy = O
End Function

Function AywLik(A, Lik) As String()
Dim I
For Each I In Itr(A)
    If I Like Lik Then PushI AywLik, I
Next
End Function

Function AywLikAy(A, LikAy$()) As String()
Dim I, Lik
For Each I In Itr(A)
    For Each Lik In LikAy
        If I Like Lik Then
            PushI AywLikAy, I
            Exit For
        End If
    Next
Next
End Function
Function IsEmpWhNm(A As WhNm) As Boolean
IsEmpWhNm = True
If IsNothing(A) Then Exit Function
With A
    If IsNothing(.Re) Then
        If Sz(.ExlLikAy) = 0 Then
            If Sz(.LikAy) = 0 Then
                Exit Function
            End If
        End If
    End If
End With
IsEmpWhNm = True
End Function
Function AywWhStrPfx(A, WhStr$, Optional NmPfx$) As String()
AywWhStrPfx = AywNm(A, WhNmzStr(WhStr, NmPfx))
End Function

Function AywNmStr(A, WhStr$, Optional NmPfx$) As String()
AywNmStr = AywNm(A, WhNmzStr(WhStr, NmPfx))
End Function

Function AywNm(A, B As WhNm) As String()
Dim I
For Each I In Itr(A)
    If HitNm(I, B) Then PushI AywNm, I
Next
End Function

Function AyePfx(A, Pfx$) As String()
Dim I
For Each I In Itr(A)
    If Not HasPfx(I, Pfx) Then PushI AyePfx, I
Next
End Function

Function AywObjPred(A, Obj, Pred$)
Dim I, O, X
AywObjPred = AyCln(A)
For Each I In Itr(A)
    X = CallByName(Obj, Pred, VbMethod, I)
    If X Then
        Push AywObjPred, I
    End If
Next
End Function

Function AywPatn(A, Patn$) As String()
If Sz(A) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AywPatn = SyzAy(A): Exit Function
Dim X, R As RegExp
Set R = RegExp(Patn)
For Each X In Itr(A)
    If R.Test(X) Then Push AywPatn, X
Next
End Function

Function AywPatnExl(A, Patn$, ExlLikss$) As String()
AywPatnExl = AyeLikss(AywPatn(A, Patn), ExlLikss)
End Function

Function AyPatn_IxAy(A, Patn$) As Long()
AyPatn_IxAy = AyRe_IxAy(A, RegExp(Patn))
End Function
Function AyRe_IxAy(A, B As RegExp) As Long()
If Sz(A) = 0 Then Exit Function
Dim I, O&(), J&
For Each I In A
    If B.Test(I) Then Push O, J
    J = J + 1
Next
AyRe_IxAy = O
End Function

Function AywPfx(A, Pfx$) As String()
Dim I
For Each I In Itr(A)
    If HasPfx(I, Pfx) Then PushI AywPfx, I
Next
End Function

Function AywPred(A, Pred$)
Dim X
AywPred = AyCln(A)
For Each X In Itr(A)
    If Run(Pred, X) Then
        Push AywPred, X
    End If
Next
End Function

Function AywPredFalse(A, Pred$)
Dim X
AywPredFalse = AyCln(A)
For Each X In Itr(A)
    If Not Run(Pred, X) Then
        Push AywPredFalse, X
    End If
Next
End Function

Function AywPredNot(A, Pred$)
AywPredNot = AywPredFalse(A, Pred)
End Function

Function AywPredXAB(Ay, XAB$, A, B)
Dim X
AywPredXAB = AyCln(Ay)
For Each X In Itr(Ay)
    If Run(XAB, X, A, B) Then
        Push AywPredXAB, X
    End If
Next
End Function

Function AywPredXABC(Ay, XABC$, A, B, C)
Dim X
AywPredXABC = AyCln(Ay)
For Each X In Itr(Ay)
    If Run(XABC, X, A, B, C) Then
        Push AywPredXABC, X
    End If
Next
End Function

Function AywPredXAP(A, PredXAP$, ParamArray Ap())
AywPredXAP = AyCln(A)
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In Itr(A)
    Asg I, Av(0)
    If RunAv(PredXAP, Av) Then
        Push AywPredXAP, I
    End If
Next
End Function

Function AywPredXP(A, XP$, P)
Dim X
AywPredXP = AyCln(A)
For Each X In Itr(A)
    If Run(XP, X, P) Then
        Push AywPredXP, X
    End If
Next
End Function

Function AywPredXPNot(A, XP$, P)
Dim X
AywPredXPNot = AyCln(A)
For Each X In Itr(A)
    If Not Run(XP, X, P) Then
        Push AywPredXPNot, X
    End If
Next
End Function

Function AywRe(A, Re As RegExp) As String()
If IsNothing(Re) Then AywRe = SyzAy(A): Exit Function
Dim X
For Each X In Itr(A)
    If Re.Test(X) Then PushI AywRe, X
Next
End Function
Function AywRmvEle(A, Ele)
AywRmvEle = AyCln(A)
Dim I
For Each I In Itr(A)
    If I <> Ele Then PushI AywRmvEle, I
Next
End Function
Function ItrzAywRmvT1(A, T1$)
Asg Itr(AywRmvT1(A, T1)), ItrzAywRmvT1
End Function

Function ItrzRmvT1(Ay, T1$)
Asg Itr(AywRmvT1(Ay, T1)), ItrzRmvT1
End Function

Function AywRmvT1(Ay, T1$) As String()
AywRmvT1 = AyRmvT1(AywT1(Ay, T1))
End Function

Function AywRmvTT(A, T1$, T2$) As String()
AywRmvTT = AyRmvTT(AywTT(A, T1, T2))
End Function

Function AywSfx(A, Sfx$) As String()
Dim I
For Each I In Itr(A)
    If HasSfx(I, Sfx) Then PushI AywSfx, I
Next
End Function

Function AywSingleEle(A)
Dim O: O = A: Erase O
Dim CntDry(): CntDry = CntDryzAy(A)
If Sz(CntDry) = 0 Then
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

Function AywSng(A)
AywSng = AyMinus(A, AywDup(A))
End Function

Function AywSngEle(A)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = AyCln(A)
Dim K, D As Dictionary
Set D = CntDic(A)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AywT1(Ay, T1) As String()
Dim L
For Each L In Itr(Ay)
    If HasT1(L, T1) Then
        PushI AywT1, L
    End If
Next
End Function

Function AywT1InAy(A, Ay$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), L
For Each L In A
    If HasEle(Ay, T1(L)) Then Push O, L
Next
AywT1InAy = O
End Function

Function AywT1SelRst(A, T1) As String()
Dim L
For Each L In Itr(A)
    If ShfTerm(L) = T1 Then PushI AywT1SelRst, L
Next
End Function

Function AywT2EqV(A$(), V) As String()
AywT2EqV = AywPredXP(A, "HasL_T2", V)
End Function

Function AywTT(A, T1$, T2$) As String()
AywTT = AywPredXAB(A, "HasTT", T1, T2)
End Function

Function AywTTSelRst(A, T1, T2) As String()
Dim L, X1$, X2$, Rst$
For Each L In Itr(A)
    Asg2TRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI AywTTSelRst, Rst
        End If
    End If
Next
End Function

Function SywFT(A$(), FmIx, ToIx) As String()
Dim J&
For J = FmIx To ToIx
    Push SywFT, A(J)
Next
End Function

Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C As Boolean
Dim D&
Dim E As FTIx
Dim F$()
Dim G As WhNm
Dim H()
Dim I As RegExp
AyDupAss A, B
AywDist A, C
AywDistFmt A
AywDistSy A
AywDistT1 A
AywDup A
AywFTIx A, E
AywRmvT1 A, B
AywRmvTT A, B, B
AywSfx A, B
AywSingleEle A
AywSng A
AywSngEle A
AywT1 A, A
AywT1InAy A, F
AywT1SelRst A, A
AywT2EqV F, A
AywTT A, B, B
AywTTSelRst A, A, A
SywFT F, A, A
End Sub

Private Sub Z()
End Sub
