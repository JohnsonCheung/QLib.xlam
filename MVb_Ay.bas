Attribute VB_Name = "MVb_Ay"
Option Explicit
Public Const DoczStmt$ = "Stmt is a string between StmtBrkPatn"
Public Const StmtBrkPatn$ = "(\.  |\r\n|\r)"

Sub AsgAp(Ay, ParamArray OAp())
Dim J%, OAv()
OAv = OAp
For J = 0 To Min(UB(Ay), UB(OAv))
    OAp(J) = Ay(J)
Next
End Sub

Sub AsgT1SyRestSy(Sy$(), OT1Sy$(), ORestSy$())
OT1Sy = T1Sy(Sy)
ORestSy = SyRmvT1(Sy)
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

Function AyCln(Ay)
ThwIfNotAy Ay, CSub
Dim O
O = Ay
Erase O
AyCln = O
End Function

Function ChkDup(Ay, QMsg$) As String()
Dim Dup
Dup = AywDup(Ay)
If Si(Dup) = 0 Then Exit Function
PushI ChkDup, FmtQQ(QMsg, JnSpc(Dup))
End Function

Function DupT1(Sy$()) As String()
DupT1 = CvSy(AywDup(T1Sy(Sy)))
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
If IsEmp(O1) Then Exit Function
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
O = AyCln(AyOfAy(0))
Dim X
For Each X In AyOfAy
    PushAy O, X
Next
AyOfAyAy = O
End Function

Private Sub Z_AyFlat()
Dim AyOfAy()
AyOfAy = Array(SySsl("a b c d"), SySsl("a b c"))
Ept = SySsl("a b c d a b c")
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

Function LasEle(Ay)
Dim N%: N = Si(Ay)
If N = 0 Then
    Warn CSub, "No ele in Ay"
Else
    Asg Ay(N - 1), LasEle
End If
End Function

Function AyMid(Ay, Fm, Optional L = 0)
AyMid = AyCln(Ay)
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

Function NxtFn$(Fn$, FnSy$(), Optional MaxN% = 999)
If Not HasEle(FnSy, Fn) Then NxtFn = Fn: Exit Function
NxtFn = MaxzAy(SywLik(FnSy, Fn & "(???)"))
End Function

Function ItrzLines(Lines$)
Asg Itr(SplitCrLf(Lines$)), ItrzLines
End Function

Function Itr(Ay)
If Si(Ay) = 0 Then Set Itr = New Collection Else Itr = Ay
End Function

Function AyRTrim(Sy$()) As String()
If Si(Sy) = 0 Then Exit Function
Dim O$(), I
For Each I In Sy
    Push O, RTrim(I)
Next
AyRTrim = O
End Function
Sub ResiN(OAy, N&)
Resi OAy, N - 1
End Sub

Sub Resi(OAy, U&)
If U < 0 Then Erase OAy: Exit Sub
ReDim OAy(U)
End Sub

Function ReserveAy(Ay)
Dim O: O = Ay
Dim J&, U&
U = UB(O)
For J = 0 To U
    Asg Ay(U - J), O(J)
Next
ReserveAy = O
End Function

Function ReverseAyI(Ay)
Dim O: O = Ay
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = Ay(U - J)
Next
ReverseAyI = O
End Function

Function ReverseOy(Oy() As Object)
Dim O: O = Oy
Dim J&, U&
U = UB(O)
For J = 0 To U
    Set O(J) = Oy(U - J)
Next
ReverseOy = O
End Function

Function SyRplMid(Ay, B As FTIx, ByAy)
Dim M As AyABC: Set M = AyabcByFTIx(Ay, B)
SyRplMid = AyAddAp(M.A, ByAy, M.C)
End Function

Function AySampLin$(Ay)
Dim S$, U&
U = UB(Ay)
If U >= 0 Then
    Select Case True
    Case IsPrim(Ay(0)): S = "[" & Ay(0) & "]"
    Case IsObject(Ay(0)), IsArray(Ay(0)): S = "[*Ty:" & TypeName(Ay(0)) & "]"
    Case Else: Stop
    End Select
End If
AySampLin = "*Ay:[" & U & "]" & S
End Function

Function Itm_IsSel(Itm, Ay) As Boolean
If Si(Ay) = 0 Then Itm_IsSel = True: Exit Function
Itm_IsSel = HasEle(Ay, Itm)
End Function

Function SeqCntDicvAy(Ay) As Dictionary 'The return dic of key=AyEle pointing to 2-Ele-LngAy with Ele-0 as Seq#(0..) and Ele- as Cnt
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
Set SeqCntDicvAy = O
End Function
Function SqzAyH(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To 1, 1 To N)
For Each V In Ay
    J = J + 1
    O(1, J) = V
Next
SqzAyH = O
End Function

Function SqzAyV(Ay) As Variant()
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Dim J&, V
Dim O()
ReDim O(1 To N, 1 To 1)
For Each V In Ay
    J = J + 1
    O(J, 1) = V
Next
SqzAyV = O
End Function

Function AywT1SelRst(Sy$(), T1$) As String()
AywT1SelRst = SyRmvT1(AywT1(Sy, T1))
End Function

Function AyIndent(Sy$(), Optional Ident% = 4) As String()
Dim I, S$
S = Space(Ident)
For Each I In Itr(Sy)
    PushI AyIndent, S & I
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
Function WdtzSy%(Sy$())
Dim O%, J&
For J = 0 To UB(Sy)
    O = Max(O, Len(Sy(J)))
Next
WdtzSy = O
End Function

Function SyWrpPad(Sy$(), W%) As String() ' Each Itm of Sy[Sy] is padded to line with WdtzSy(Sy).  return all padded lines as String()
Dim O$(), X, I%
ReDim O(0)
For Each X In Itr(Sy)
    If Len(O(I)) + Len(X) < W Then
        O(I) = O(I) & X
    Else
        PushI O, X
        I = I + 1
    End If
Next
SyWrpPad = O
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
Function StmtLy(StmtLin$) As String()
StmtLy = SyEnsSfxDot(AyLTrim(Split(StmtLin, ". ")))
End Function
Function AyZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O()
    Dim J&
    Resi O, U
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
    Resi ODry, URow
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

Function StrAddSy(S$, Sy$()) As String()
StrAddSy = CvSy(ItmAddAy(S, Sy))
End Function


Function ItmAddAy(Itm, Ay)
ItmAddAy = AyInsEle(Ay, Itm)
End Function

Private Sub ZZZ_AyabcByFmTo()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim Act As AyABC: Act = AyabcByFmTo(Ay, 1, 2)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Private Sub ZZ_AyAsgAp()
Dim O%, Ay$
'AyAsgAp Array(234, "abc"), O, Ay
Ass O = 234
Ass Ay = "abc"
End Sub

Private Sub ZZ_ChkEqAy()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub ZZ_MaxAy()
Dim Ay()
Dim Act
Act = MaxAy(Ay)
Stop
End Sub

Private Sub ZZ_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ThwIfNE Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ThwIfNE Exp, Act
End Sub

Private Sub ZZ_AyeEmpEleAtEnd()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub ZZ_SyzAy()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Si(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub ZZ_AyTrim()
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

Private Sub Z_ChkEqAy()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub Z_AyabcByFTIxIx()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim M As FTIx: M = FTIx(1, 2)
Dim Act As AyABC: Act = AyabcByFTIx(Ay, M)
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

Private Sub Z_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
ThwAyabNE Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
ThwAyabNE Exp, Act
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


Private Sub Z()
Z_AyFlat
Z_AyabcByFTIxIx
Z_HasEleDupEle
Z_ChkEqAy
Z_AyMinus
Z_SyzAy
Z_AyTrim
MVb_Ay:
End Sub

Private Sub Z_AddPfxToSsl()
Dim Ssl$, Exp$(), Pfx$
Ssl = "B C D"
Pfx = "A"
Exp = SySsl("AB AC AD")
GoSub Tst
Exit Sub
Tst:
    Dim Act$()
    Act = AddPfxToSsl(Pfx, Ssl)
    Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function AddPfxToSsl(Pfx$, Ssl$) As String()
AddPfxToSsl = SyAddPfx(SySsl(Ssl), Pfx)
End Function

Function ItrzSsl(Ssl$)
Asg Itr(SySsl(Ssl)), ItrzSsl
End Function
Function SySsl(S$) As String()
SySsl = SplitSpc(Trim(RplDblSpc(S)))
End Function

Function IntSeq(N&, Optional IsFmOne As Boolean) As Integer()
Dim O%(): ReDim O(N - 1)
Dim J&
If IsFmOne Then
    For J = 0 To N - 1
        O(J) = J + 1
    Next
Else
    For J = 0 To N - 1
        O(J) = J
    Next
End If
IntSeq = O
End Function

