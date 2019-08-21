Attribute VB_Name = "QVb_Ay_SubSet_Aw"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Sub_Wh."
Private Const Asm$ = "QVb"

Sub ThwIf_Dup(Ay, Fun$)
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Sub
Thw Fun, "There are dup in array", "Dup Ay", Dup, Ay
End Sub
Function AwIxCnt(Ay, Ix, Cnt)
Dim J&
Dim O: O = Ay: Erase O
For J = 0 To Cnt - 1
    Push O, Ay(Ix + J)
Next
AwIxCnt = O
End Function

Function AwBet(Ay, FmEle, ToEle)
Dim O: O = ResiU(Ay)
Dim I: For Each I In Itr(Ay)
    If IsBet(I, FmEle, ToEle) Then
        Push O, I
    End If
Next
AwBet = O
End Function

Function AwDistAsSy(Ay, Optional IgnCas As Boolean) As String()
AwDistAsSy = CvSy(AwDist(Ay, IgnCas))
End Function

Function AwDistAsI(Ay, Optional IgnCas As Boolean) As Integer()
AwDistAsI = CvIntAy(AwDist(Ay, IgnCas))
End Function

Function AwDist(Ay, Optional IgnCas As Boolean)
AwDist = IntozItr(ResiU(Ay), DiKqCnt(Ay).Keys)
End Function

Private Sub Z_FmtDiKqCnt()
Dim Ay
GoSub Z
Exit Sub
Z:
    Ay = Array(1, 2, 2, 2, 3, "skldflskdfsdklf" & vbCrLf & "ksdlfj")
    Brw FmtDiKqCnt(Ay)
    
End Sub


Function AwDistT1(Ay) As String()
AwDistT1 = AwDist(T1Ay(Ay))
End Function

Function AwDup(Ay, Optional C As VbCompareMethod = vbTextCompare)
Dim O: O = ResiU(Ay)
Dim D As Dictionary: Set D = DiKqCnt(Ay, EiCntDup, C)
AwDup = IntozItr(O, D.Keys)
End Function
Function AwNonEmp(Ay)
AwNonEmp = Ay: Erase AwNonEmp
Dim I
For Each I In Ay
    If Not IsEmpty(I) Then
        PushI AwNonEmp, I
    End If
Next
End Function

Function AwFE(Ay, FmIx, EIx)
AwFE = AwFT(Ay, FmIx, EIx - 1)
End Function

Function AwFT(Ay, FmIx, ToIx)
Dim J&, I&
Dim O: O = ResiU(Ay, ToIx - FmIx)
For J = FmIx To ToIx
    Asg Ay(J), O(I)
    I = I + 1
Next
AwFT = O
End Function

Function AwFm(Ay, FmIx)
Dim J&, I&
Dim O: O = ResiU(Ay, UB(Ay) - FmIx)
For J = FmIx To UB(Ay)
    Asg Ay(J), O(I)
    I = I + 1
Next
AwFm = O
End Function

Function AwFstUEle(Ay, U)
If U > UB(Ay) Then AwFstUEle = Ay: Exit Function
Dim O: O = Ay
ReDim Preserve O(U)
AwFstUEle = O
End Function

Function AwFei(Ay, B As Fei)
AwFei = AwFE(Ay, B.FmIx, B.EIx)
End Function

Function AwIxyzMust(Ay, Ixy&())
If IsIxyOut(Ixy, UB(Ay)) Then Thw CSub, "Some element in Ixy is outsize Ay", "UB(Ay) Ixy", UB(Ay), Ixy
Dim U&: U = UB(Ay)
Dim O: O = ResiU(Ay)
Dim Ix
For Each Ix In Itr(Ixy)
    Push O, Ay(Ix)
Next
AwIxyzMust = O
End Function

Function AwInAset(Ay, B As Aset)
AwInAset = ResiU(Ay)
Dim I
For Each I In Itr(Ay)
    If Ay.Has(I) Then Push AwInAset, I
Next
End Function

Function AwIxy(Ay, Ixy&())
If Si(Ixy) = 0 Then AwIxy = Ay: Erase AwIxy: Exit Function
Dim U&: U = UB(Ixy)
Dim O: O = ResiU(Ay)
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
AwIxy = O
End Function

Function AwIxyAlwE(Ay, Ixy&())
Dim U&: U = UB(Ixy)
Dim O: O = ResiU(Ay)
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
AwIxyAlwE = O
End Function

Function AwLik(Ay, Lik) As String()
Dim I: For Each I In Itr(Ay)
    If I Like Lik Then PushI AwLik, I
Next
End Function

Function AwKssAy(Ay, KssAy$()) As String()
Dim LikAy$(): LikAy = TermAsetzTLiny(KssAy).Sy
Dim S: For Each S In Itr(Ay)
    If HitLikAy(S, LikAy) Then PushI AwKssAy, S
Next
End Function

Function AwLikAy(Ay, LikAy$()) As String()
Dim I, Lik
For Each I In Itr(Ay)
    If HitLikAy(I, LikAy) Then PushI AwLikAy, I
Next
End Function

Function AwNmStr(Ay, WhNmStr$) As String()
AwNmStr = AwNm(Ay, WhNmzS(WhNmStr))
End Function

Function AwIsNm(Ay) As String()
AwIsNm = AwPred(Ay, PredIsNm)
End Function

Function AwNm(Ay, B As WhNm) As String()
Dim I
For Each I In Itr(Ay)
    If HitNm(I, B) Then PushI AwNm, I
Next
End Function


Function AwPred(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If P.Pred(I) Then
        PushI AwPred, I
    End If
Next
End Function

Function PatnPred(Patn$) As IPred
Dim O As New PredPatn
O.Init Patn
Set PatnPred = O
End Function

Function AwPatn1(Ay, Patn$) As Variant()
If Si(Ay) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AwPatn1 = Ay: Exit Function
Dim Re As RegExp: Set Re = RegExp(Patn)
Dim I: For Each I In Itr(Ay)
    If IsStr(I) Then
        If Re.Test(I) Then PushI AwPatn1, I
    End If
Next
End Function

Function AwPatn(Ay, Patn$) As String()
If Si(Ay) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AwPatn = Ay: Exit Function
AwPatn = AwPred(Ay, PatnPred(Patn))
End Function
Function AwPatnAy(Ay, PatnAy$()) As Variant()
If Si(Ay) = 0 Then Exit Function
Stop
End Function
Function AwPatnExl(Ay, Patn$, ExlKss$) As String()
AwPatnExl = AeKss(AwPatn(Ay, Patn), ExlKss)
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

Function AwPredFalse(Ay, P As IPred)
Dim X
AwPredFalse = ResiU(Ay)
For Each X In Itr(Ay)
    If Not P.Pred(X) Then
        Push AwPredFalse, X
    End If
Next
End Function

Function AwPredXAP(Ay, PredXAP$, ParamArray Ap())
AwPredXAP = ResiU(Ay)
Dim I
Dim Av()
    Av = Ap
    Av = InsEle(Av)
For Each I In Itr(Ay)
    Asg I, Av(0)
    If RunAv(PredXAP, Av) Then
        Push AwPredXAP, I
    End If
Next
End Function

Function AwPredXP(Ay, Xp$, P)
Dim X
AwPredXP = ResiU(Ay)
For Each X In Itr(Ay)
    If Run(Xp, X, P) Then
        Push AwPredXP, X
    End If
Next
End Function

Function AwPredXPNot(Ay, Xp$, P)
Dim X
AwPredXPNot = ResiU(Ay)
For Each X In Itr(Ay)
    If Not Run(Xp, X, P) Then
        Push AwPredXPNot, X
    End If
Next
End Function

Function AwRe(Ay, Re As RegExp) As String()
Dim S
For Each S In Itr(Ay)
    If Re.Test(S) Then PushI AwRe, S
Next
End Function
Function AwRmvEle(Ay, Ele)
AwRmvEle = ResiU(Ay)
Dim I
For Each I In Itr(Ay)
    If I <> Ele Then PushI AwRmvEle, I
Next
End Function
Function ItrzAwRmvT1(Ay, T1)
Asg Itr(AwRmvT1(Ay, T1)), ItrzAwRmvT1
End Function

Function AwRmvT1(Ay, T1) As String()
AwRmvT1 = RmvT1zAy(AwT1(Ay, T1))
End Function

Function AwRmvTT(Ay, T1, T2) As String()
AwRmvTT = RmvTTzAy(AwTT(Ay, T1, T2))
End Function

Function AwSkip(Ay, Optional SkipN& = 1)
If SkipN <= 0 Then AwSkip = Ay: Exit Function
Dim U&: U = UB(Ay) - SkipN: If SkipN < -1 Then Thw CSub, "Ay is not enough to skip", "Si-Ay SkipN", "Si(Ay),SKipN"
Dim O: O = ResiU(Ay, U)
Dim J&, I&
For J = SkipN To UB(Ay)
    Asg Ay(J), O(I)
    I = I + 1
Next
End Function

Function AwSfx(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    If HasSfx(I, Sfx) Then PushI AwSfx, I
Next
End Function

Function AwSingleEle(Ay)
Dim O: O = Ay: Erase O
Dim CntDy(): CntDy = CntDyoAy(Ay)
If Si(CntDy) = 0 Then
    AwSingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDy
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
AwSingleEle = O
End Function

Function AwSng(Ay)
AwSng = MinusAy(Ay, AwDup(Ay))
End Function

Function AwSngEle(Ay)
'Return Set of Element as array in {Ay} having 2 or more element
Dim O: O = ResiU(Ay)
Dim K, D As Dictionary
Set D = DiKqCnt(Ay)
For Each K In D.Keys
    If D(K) = 1 Then PushI O, K
Next
End Function

Function AwT1(Ay, T1) As String()
Dim L$, I
For Each I In Itr(Ay)
    L = I
    If HasT1(L, T1) Then
        PushI AwT1, L
    End If
Next
End Function

Function AwT1InAy(Ay, InAy) As String()
If Si(Ay) = 0 Then Exit Function
Dim O$(), L: For Each L In Ay
    If HasEle(InAy, T1(L)) Then Push O, L
Next
AwT1InAy = O
End Function

Function AwT1SelRst(Ay, T1) As String()
Dim I, L$
For Each I In Itr(Ay)
    L = I
    If ShfT1(L) = T1 Then PushI AwT1SelRst, L
Next
End Function

Function AwTT(Ay, T1, T2) As String()
Dim I, L$
For Each I In Itr(Ay)
    L = I
    If HasTT(L, T1, T2) Then PushI AwTT, L
Next
End Function

Function AwTTSelRst(Ay, T1, T2) As String()
Dim L$, I, X1$, X2$, Rst$
For Each I In Itr(Ay)
    L = I
    AsgTTRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            PushI AwTTSelRst, Rst
        End If
    End If
Next
End Function

Private Sub Z()
Dim Ay As Variant
Dim B$
Dim C As Boolean
Dim D&
Dim E As Fei
Dim F$()
Dim G As WhNm
Dim H()
Dim I As RegExp
AwDist Ay, C
FmtDiKqCnt Ay
End Sub

