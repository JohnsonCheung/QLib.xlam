Attribute VB_Name = "QVb_Ay"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay."
Private Const Asm$ = "QVb"
Public Const DoczStmt$ = "Stmt is a string between StmtBrkPatn"
Public Const StmtBrkPatn$ = "(\.  |\r\n|\r)"
Public Const DoczSsLin = "Spc-Sep-Lin.  SplitSpc after trim and rmvDblSpc."
Public Const DoczSsl$ = "Spc-Sep-Lin-Escaped.  SpcSepStrRev each ele after SyzSS"

Sub AsgAp(Ay, ParamArray OAp())
Dim J%, OAv()
OAv = OAp
For J = 0 To Min(UB(Ay), UB(OAv))
    OAp(J) = Ay(J)
Next
End Sub

Sub AsgT1SyRestSy(Sy$(), OT1Sy$(), ORestSy$())
OT1Sy = T1Ay(Sy)
ORestSy = RmvT1zAy(Sy)
End Sub

Function VcAy(Ay, Optional Fnn$)
BrwAy Ay, Fnn, UseVc:=True
End Function

Function BrwAy(Ay, Optional Fnn$, Optional UseVc As Boolean)
Dim T$
T = TmpFt("BrwAy", Fnn)
WrtAy Ay, T
BrwFt T, UseVc
BrwAy = Ay
End Function


Function ChkDup(Ay, QMsg$) As String()
Dim Dup
Dup = AywDup(Ay)
If Si(Dup) = 0 Then Exit Function
PushI ChkDup, FmtQQ(QMsg, JnSpc(Dup))
End Function
Function LyzVbl(Vbl) As String()
LyzVbl = SplitVBar(Vbl)
End Function
Function DupT1Ay(Ly$(), Optional C As VbCompareMethod = vbTextCompare) As String()
Dim A$(): A = T1Ay(Ly)
DupT1Ay = AywDup(A, C)
End Function

Function ChkAyEmp(Ay, Msg$) As String()
If Si(Ay) = 0 Then ChkAyEmp = Sy(Msg)
End Function

Function ChkEqAy(Ay1, Ay2, Optional N1$ = "Exp", Optional N2$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Si: [?] [?]", N1, N2, Si(Ay1), Si(Ay2)): GoTo X
If Si(Ay1) = 0 Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", N1, N2, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
'If IsEmp(O1) Then Exit Function
Dim O2$()
    Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", N1, N2, Si(Ay1))
    If ReachLimit Then
        Push O2, FmtQQ("At least [?] differences:", Si(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, FmtQQ("Ay-[?]:", N1)
PushIAy O, SyQuote(SyzAy(Ay1), "[]")
Push O, FmtQQ("Ay-[?]:", N2)
PushIAy O, SyQuote(SyzAy(Ay2), "[]")
ChkEqAy = O
End Function

Function AyOfAyAy(AyOfAy)
If Si(AyOfAy) = 0 Then Exit Function
Dim O
O = Resi(AyOfAy(0))
Dim X
For Each X In AyOfAy
    PushAy O, X
Next
AyOfAyAy = O
End Function

Private Sub Z_AyFlat()
Dim AyOfAy()
AyOfAy = Array(SyzSS("a b c d"), SyzSS("a b c"))
Ept = SyzSS("a b c d a b c")
GoSub Tst
Exit Sub
Tst:
    Act = AyFlat(AyOfAy)
    C
    Return
End Sub

Function AyFlat(AyOfAy)
AyFlat = AyOfAyAy(AyOfAy)
End Function

Function AyItmCnt%(Ay, M)
If Si(Ay) = 0 Then Exit Function
Dim O%, X
For Each X In Itr(Ay)
    If X = M Then O = O + 1
Next
AyItmCnt = O
End Function
Function AywSubStr(Ay, SubStr) As String()
AywSubStr = AywPred(Ay, PredzSubStr(SubStr))
End Function
Function AywPredzSy(Ay, P As IPred) As String()
Dim I
For Each I In Itr(Ay)
    If P.Pred(I) Then PushI AywPredzSy, I
Next
End Function

Function AywPfx(Ay, Pfx) As String()
AywPfx = AywPred(Ay, PredzPfx(Pfx))
End Function

Function AywLasN(Ay, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(Ay)
If U < N Then AywLasN = Ay: Exit Function
O = Ay
Fm = U - N + 1
NewU = N - 1
For J = Fm To U
    Asg O(J), O(I)
    I = I + 1
Next
ReDim Preserve O(NewU)
AywLasN = O
End Function
Function LasSndEle(Ay)
Dim N&: N = Si(Ay)
If N <= 1 Then
    Thw CSub, "Only 1 or no ele in Ay"
Else
    Asg Ay(N - 2), LasSndEle
End If
End Function
Function LasEle(Ay)
Dim N&: N = Si(Ay)
If N = 0 Then
    Thw CSub, "No ele in Ay"
Else
    Asg Ay(N - 1), LasEle
End If
End Function

Function AyMid(Ay, Fm, Optional L = 0)
AyMid = Resi(Ay)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(Ay)
    Case Else:  E = Min(UB(Ay), L + Fm - 1)
    End Select
For J = Fm To E
    Push AyMid, Ay(J)
Next
End Function

Function NxtNm$(Ny$(), Optional MaxN% = 0)
Stop
End Function

Function NxtFn$(Fn$, FnAy$(), Optional MaxN% = 999)
If Not HasEle(FnAy, Fn) Then NxtFn = Fn: Exit Function
NxtFn = MaxzAy(AywLik(FnAy, Fn & "(???)"))
End Function

Function ItrzLines(Lines$)
Asg Itr(SplitCrLf(Lines$)), ItrzLines
End Function

Function NItr&(Itr)
Dim O&, V
For Each V In Itr
    O = O + 1
Next
NItr = O
End Function

Function Itr(Ay)
If Si(Ay) = 0 Then Set Itr = New Collection Else Itr = Ay
End Function

Function RTrimAy(Ay) As String()
If Si(Ay) = 0 Then Exit Function
Dim O$(), I
For Each I In Ay
    Push O, RTrim(I)
Next
RTrimAy = O
End Function

Function ResiN(Ay, N&)
ResiN = Resi(Ay, N - 1)
End Function

Function ResiMax(OAy1, OAy2) 'Resi the smaller Ay to same si as the larger Ay and return fst Ay
Dim U1&, U2&: U1 = UB(OAy1): U2 = UB(OAy2)
Select Case True
Case U1 > U2: OAy2 = Resi(OAy2, U1)
Case U2 > U1: OAy1 = Resi(OAy1, U2)
End Select
ResiMax = OAy1
End Function

Function Resi(Ay, Optional U& = -1) 'Return the resized Ay
Dim O: O = Ay
If U < 0 Then Erase O: Resi = O: Exit Function
ReDim Preserve O(U)
Resi = O
End Function

Function RevAy(Ay) 'Return reversed Ay
Dim O: O = Ay
Dim J&, U&
U = UB(O)
For J = 0 To U
    Asg Ay(U - J), O(J)
Next
RevAy = O
End Function

Function RevAyAyI(Ay)
Dim O: O = Ay
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = Ay(U - J)
Next
RevAyAyI = O
End Function

Function RevAyOy(Oy() As Object)
Dim O: O = Oy
Dim J&, U&
U = UB(O)
For J = 0 To U
    Set O(J) = Oy(U - J)
Next
RevAyOy = O
End Function

Function RplAyzMid(Ay, B As Fei, ByAy)
With AyabczAyFei(Ay, B)
RplAyzMid = AddAyAp(.A, ByAy, .C)
End With
End Function

Function SampLinzAy$(Ay)
Dim S$, U&
U = UB(Ay)
If U >= 0 Then
    Select Case True
    Case IsPrim(Ay(0)): S = "[" & Ay(0) & "]"
    Case IsObject(Ay(0)), IsArray(Ay(0)): S = "[*Ty:" & TypeName(Ay(0)) & "]"
    Case Else: Stop
    End Select
End If
SampLinzAy = "*Ay:[" & U & "]" & S
End Function

Function SeqCntDic(Ay) As Dictionary 'The return dic of key=AyEle pointing to 2-Ele-LngAp with Ele-0 as Seq#(0..) and Ele- as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In Itr(Ay)
    If O.Exists(X) Then
        L = O(X)
        L(1) = L(1) + 1
        O(X) = L
    Else
        ReDim L(1)
        L(0) = S
        L(1) = 1
        O.Add X, L
    End If
Next
Set SeqCntDic = O
End Function
Function StrColzSq(Sq(), Optional C = 1) As String()
If Si(Sq) = 0 Then Exit Function
Dim R&
For R = 1 To UBound(Sq, 1)
    PushI StrColzSq, Sq(R, C)
Next
End Function
Function SqhzAp(ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
SqhzAp = Sqh(Av)
End Function

Function Sqh(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To 1, 1 To N)
For Each V In Ay
    J = J + 1
    O(1, J) = V
Next
Sqh = O
End Function

Function Sqv(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To N, 1 To 1)
For Each V In Ay
    J = J + 1
    O(J, 1) = V
Next
Sqv = O
End Function

Function AywT1SelRst(Sy$(), T1) As String()
AywT1SelRst = RmvT1zAy(AywT1(Sy, T1))
End Function

Function IndentSy(Sy$(), Optional Indent% = 4) As String()
Dim I, S$
S = Space(Indent)
For Each I In Itr(Sy)
    PushI IndentSy, S & I
Next
End Function
Function AyTrim(Sy$()) As String()
Dim X
For Each X In Itr(Sy)
    Push AyTrim, Trim(X)
Next
End Function
Function MinzAy(Ay)
Dim O, I
For Each I In Ay
    If I < O Then O = I
Next
MinzAy = O
End Function
Function MaxzAy(Ay)
Dim O, I
For Each I In Ay
    If I > O Then O = I
Next
MaxzAy = O
End Function
Function WdtzAy%(Ay)
Dim O%, V
For Each V In Itr(Ay)
    O = Max(O, Len(V))
Next
WdtzAy = O
End Function

Sub WrtAy(Ay, Ft$, Optional OvrWrt As Boolean)
WrtStr JnCrLf(Ay), Ft, OvrWrt
End Sub
Function AyLTrim(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI AyLTrim, LTrim(L)
Next
End Function
Function SyEnsSfxDot(Ay) As String()
SyEnsSfxDot = SyEnsSfx(Sy, ".")
End Function
Function SyEnsSfx(Sy$(), Sfx$) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyEnsSfx, EnsSfx(CStr(I), Sfx)
Next
End Function
Function StmtLy(StmtLin) As String()
StmtLy = SyEnsSfxDot(AyLTrim(Split(StmtLin, ". ")))
End Function
Function AyZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O()
    Dim J&
    O = Resi(O, U)
    For J = 0 To U
        If U1 >= J Then
            If U2 >= J Then
                O(J) = Array(A1(J), A2(J))
            Else
                O(J) = Array(A1(J), Empty)
            End If
        Else
            If U2 >= J Then
                O(J) = Array(, A2(J))
            Else
                Stop
            End If
        End If
    Next
AyZip = O
End Function

Function AyZip_Ap(Ay, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
Dim UCol%
    UCol = UB(Av)

Dim URow1&
    URow1 = UB(Ay)

Dim URow&
Dim URowAy&()
    Dim J%, IURow%
    URow = URow1
    For J = 0 To UB(Av)
        IURow = UB(Av(J))
        Push URowAy, IURow
        If IURow > URow Then URow = IURow
    Next

Dim ODry()
    Dim Dr()
    ODry = Resi(ODry, URow)
    Dim I%
    For J = 0 To URow
        Erase Dr
        If URow1 >= J Then
            Push Dr, Ay(J)
        Else
            Push Dr, Empty
        End If
        For I = 0 To UB(Av)
            If URowAy(I) >= J Then
                Push Dr, Av(I)(J)
            Else
                Push Dr, Empty
            End If
        Next
        ODry(J) = Dr
    Next
AyZip_Ap = ODry
End Function

Function ItmAddAy(Itm, Ay)
ItmAddAy = AyInsEle(Ay, Itm)
End Function

Private Sub Z_AyabczAyFE()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim Act As Ayabc: Act = AyabczAyFE(Ay, 1, 2)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Private Sub Z_AyAsgAp()
Dim O%, Ay$
'AyAsgAp Array(234, "abc"), O, Ay
Ass O = 234
Ass Ay = "abc"
End Sub

Private Sub Z_ChkEqAy()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub Z_MaxEle()
Dim Ay()
Dim Act
Act = MaxEle(Ay)
Stop
End Sub

Private Sub Z_MinusAy()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = MinusAy(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ThwIf_NE Exp, Act
'
Act = MinusAyAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ThwIf_NE Exp, Act
End Sub

Private Sub Z_AyeEmpEleAtEnd()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub Z_SyzAy()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub Z_AyTrim()
DmpAy AyTrim(Sy(1, 2, 3, "  a"))
End Sub


Private Sub Z_ChkDup()
Dim Ay
Ay = Array("1", "1", "2")
Ept = Sy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = ChkDup(Ay, "This item[?] is duplicated")
    C
    Return
End Sub

Private Sub Z_ChkEqAy5()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub Z_AyabczAyFei()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim M As Fei: M = Fei(1, 2)
Dim Act As Ayabc: Act = AyabczAyFei(Ay, M)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Private Sub Z_HasEleDupEle()
Ass HasEleDupEle(Array(1, 2, 3, 4)) = False
Ass HasEleDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Private Sub Z_AyInsEle()
Dim Ay, M, At&
'
Ay = Array(1, 2, 3)
M = "X"
Ept = Array("X", 1, 2, 3)
GoSub Tst
'
Exit Sub
Tst:
    Act = AyInsEle(Ay, M, At)
    C
Return
End Sub

Private Sub Z_AyInsAy()
Dim Act, Exp, Ay(), B(), At&
Ay = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAyAt(Ay, B, At)
Ass IsEqAy(Act, Exp)
End Sub

Private Sub Z_MinusAy6()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = MinusAy(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ThwIf_AyabNE Exp, Act
'
Act = MinusAyAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ThwIf_AyabNE Exp, Act
End Sub

Private Sub Z_SyzAy2()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub Z_AyTrim2()
DmpAy AyTrim(Sy(1, 2, 3, "  a"))
End Sub

Private Sub Z_KKCMiDry()
Dim Dry(), Act As KKCntMulItmColDry, KKColIx%(), IxzAy%
PushI Dry, Array()
PushI Dry, Array()
PushI Dry, Array()
PushI Dry, Array()
PushI Dry, Array()
PushI Dry, Array()
'Ass Si(Act) = 4
'Ass IsEqAy(Act(0), Array("Ay", 3, 1, 2, 3))
'Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
'Ass IsEqAy(Act(2), Array("C", 0))
'Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub


Private Sub ZZ()
Z_AyFlat
Z_HasEleDupEle
Z_ChkEqAy
Z_MinusAy
Z_SyzAy
Z_AyTrim
MVb_Ay:
End Sub

Private Sub Z_AddPfxzSslIn()
Dim Ssl$, Exp$(), Pfx$
Ssl = "B C D"
Pfx = "A"
Exp = SyzSS("AB AC AD")
GoSub Tst
Exit Sub
Tst:
    Dim Act$()
    Act = AddPfxzSslIn(Pfx, Ssl)
    Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function AddPfxzSslIn(Pfx$, SsLin) As String()
AddPfxzSslIn = AddPfxzAy(SyzSS(SsLin), Pfx)
End Function

Function SpcSepStr$(S)
If S = "" Then SpcSepStr = ".": Exit Function
SpcSepStr = QuoteSqIf(EscSqBkt(SlashCrLf(EscBackSlash(S))))
End Function

Function RevSpcSepStr$(SpcSepStr$)
If SpcSepStr = "." Then Exit Function
RevSpcSepStr = UnTidleSpc(UnSlashTab(UnSlashCrLf(SpcSepStr)))
End Function

Function SslzDrv$(Drv)
Dim J&, U&, O$()
U = UB(Drv)
If U < 0 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = SpcSepStr(CStr(Drv(J)))
Next
SslzDrv = JnSpc(O)
End Function
Function SyzSsl(Ssl$) As String()
Dim I
For Each I In SyzSS(Ssl)
    PushI SyzSsl, RevSpcSepStr(CStr(I))
Next
End Function
Function SyzSS(SS) As String()
SyzSS = SplitSpc(RplDblSpc(Trim(SS)))
End Function

Function IntSeq(N&, Optional Fm% = 0) As Integer()
Dim O%(): ReDim O(N - 1)
Dim J&
    For J = 0 To N - 1
        O(J) = J + Fm
    Next
IntSeq = O
End Function

Function ItrzTT(TT$)
Asg Itr(TermAy(TT)), ItrzTT
End Function

