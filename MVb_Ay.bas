Attribute VB_Name = "MVb_Ay"
Option Explicit

Sub Asg_ValTo_VarVarible_and_EleOfVariantAy_and_Ap()
Dim A: A = CByte(0): A = ""
Dim B(): ReDim B(0): B(0) = CByte(0): B(0) = ""
Dim C As Byte
WAsg C
End Sub
Private Sub WAsg(ParamArray C())
C(0) = ""
End Sub
Sub AsgAp(Ay, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To Min(UB(Av), UB(Ay))
    OAp(J) = Ay(J)
Next
End Sub

Sub AyAsgT1AyRestAy(A, OT1Ay$(), ORestAy$())
Dim U&, J&
U = UB(A)
If U = -1 Then
    Erase OT1Ay, ORestAy
    Exit Sub
End If
ReDim OT1Ay(U)
ReDim ORestAy(U)
For J = 0 To U
    AsgBrk A(J), " ", OT1Ay(J), ORestAy(J)
Next
End Sub
Function VcAy(A, Optional Fnn$)
BrwAy A, Fnn, UseVc:=True
End Function

Function BrwAy(A, Optional Fnn$, Optional UseVc As Boolean)
Dim T$
T = TmpFt("BrwAy", Fnn)
WrtAy A, T
BrwFt T, UseVc
BrwAy = A
End Function

Function AyCln(A)
ThwNotAy A, CSub
Dim O
O = A
Erase O
AyCln = O
End Function

Function ChkAyDup(A, QMsg$) As String()
Dim Dup
Dup = AywDup(A)
If Sz(Dup) = 0 Then Exit Function
Push ChkAyDup, FmtQQ(QMsg, JnSpc(Dup))
End Function

Function AyDupT1(A) As String()
AyDupT1 = AywDup(AyTakT1(A))
End Function

Function AyEmpChk(A, Msg$) As String()
If Sz(A) = 0 Then AyEmpChk = Sy(Msg)
End Function

Function ChkEqAy(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
If Sz(Ay1) = 0 Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", Ay1Nm, Ay2Nm, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
If IsEmp(O1) Then Exit Function
Dim O2$()
    Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", Ay1Nm, Ay2Nm, Sz(Ay1))
    If ReachLimit Then
        Push O2, FmtQQ("At least [?] differences:", Sz(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, FmtQQ("Ay-[?]:", Ay1Nm)
PushAy O, AyQuote(Ay1, "[]")
Push O, FmtQQ("Ay-[?]:", Ay2Nm)
PushAy O, AyQuote(Ay2, "[]")
ChkEqAy = O
End Function

Function AyOfAyAy(AyOfAy)
If Sz(AyOfAy) = 0 Then Exit Function
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

Function AyItmCnt%(A, M)
If Sz(A) = 0 Then Exit Function
Dim O%, X
For Each X In Itr(A)
    If X = M Then O = O + 1
Next
AyItmCnt = O
End Function

Function AywLasN(A, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(A)
If U < N Then AywLasN = A: Exit Function
O = A
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
Dim N%: N = Sz(Ay)
If N = 0 Then
    Warn CSub, "No ele in Ay"
Else
    Asg Ay(N - 1), LasEle
End If
End Function

Function AyMid(A, Fm, Optional L = 0)
AyMid = AyCln(A)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(A)
    Case Else:  E = Min(UB(A), L + Fm - 1)
    End Select
For J = Fm To E
    Push AyMid, A(J)
Next
End Function


Function AyNPfxStar%(A)
Dim O%, X
For Each X In Itr(A)
    If FstChr(X) = "*" Then AyNPfxStar = O: Exit Function
    O = O + 1
Next
End Function
Function AyNxtNm$(A, Nm$, Optional MaxN% = 99)
If Not HasEle(A, Nm) Then AyNxtNm = Nm: Exit Function
Dim J%, O$
For J = 1 To MaxN
    O = Nm & Format(J, "00")
    If Not HasEle(A, O) Then AyNxtNm = O: Exit Function
Next
Stop
End Function

Function Itr(A)
If Sz(A) = 0 Then Set Itr = New Collection Else Itr = A
End Function

Function AyRTrim(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), I
For Each I In A
    Push O, RTrim(I)
Next
AyRTrim = O
End Function
Sub ReszAyN(OAy, N)
ReszAyU OAy, N - 1
End Sub

Sub Resz(OAy, U)
ReszAyU OAy, U
End Sub

Sub ReszAyU(OAy, U)
If U < 0 Then Erase OAy: Exit Sub
ReDim OAy(U)
End Sub
Function AyReverse(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    Asg A(U - J), O(J)
Next
AyReverse = O
End Function

Function AyReverseI(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = A(U - J)
Next
AyReverseI = O
End Function

Function OyReverse(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    Set O(J) = A(U - J)
Next
OyReverse = O
End Function

Function AyRpl_MidSeg_FT_IX(A, B As FTIx, AySeg)
AyRpl_MidSeg_FT_IX = AyRpl_MidSeg(A, B.FmIx, B.ToIx, AySeg)
End Function

Function AyRpl_MidSeg(A, FmIx&, ToIx&, ByAy)
Dim M As AyABC
    M = AyABCzAyFT(A, FmIx, ToIx)
AyRpl_MidSeg = M.A
    PushAy AyRpl_MidSeg, ByAy
    PushAy AyRpl_MidSeg, M.C
End Function

Function AyRpl_Star_InEach_Ele(A$(), By) As String()
Dim X
For Each X In Itr(A)
    PushI AyRpl_Star_InEach_Ele, Replace(X, By, "*")
Next
End Function

Function AyRpl_T1(A$(), T1$) As String()
AyRpl_T1 = AyAddPfx(AyRmvT1(A), T1 & " ")
End Function

Function AySampLin$(A)
Dim S$, U&
U = UB(A)
If U >= 0 Then
    Select Case True
    Case IsPrim(A(0)): S = "[" & A(0) & "]"
    Case IsObject(A(0)), IsArray(A(0)): S = "[*Ty:" & TypeName(A(0)) & "]"
    Case Else: Stop
    End Select
End If
AySampLin = "*Ay:[" & U & "]" & S
End Function

Function Itm_IsSel(Itm, Ay) As Boolean
If Sz(Ay) = 0 Then Itm_IsSel = True: Exit Function
Itm_IsSel = HasEle(Ay, Itm)
End Function

Function SeqCntDicvAy(A) As Dictionary 'The return dic of key=AyEle pointing to 2-Ele-LngAy with Ele-0 as Seq#(0..) and Ele- as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In Itr(A)
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
Dim N&: N = Sz(Ay)
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
Dim N&: N = Sz(Ay)
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

Function AyT1Chd(A, T1) As String()
AyT1Chd = AyRmvT1(AywT1(A, T1))
End Function

Function AyIndent(A, Optional Ident% = 4) As String()
Dim I, S$
S = Space(Ident)
For Each I In Itr(A)
    PushI AyIndent, S & I
Next
End Function
Function AyTrim(A) As String()
Dim X
For Each X In Itr(A)
    Push AyTrim, Trim(X)
Next
End Function

Function WdtzAy%(A)
Dim O%, J&
For J = 0 To UB(A)
    O = Max(O, Len(A(J)))
Next
WdtzAy = O
End Function

Function AyWrpPad(A, W%) As String() ' Each Itm of Ay[A] is padded to line with WdtzAy(A).  return all padded lines as String()
Dim O$(), X, I%
ReDim O(0)
For Each X In Itr(A)
    If Len(O(I)) + Len(X) < W Then
        O(I) = O(I) & X
    Else
        PushI O, X
        I = I + 1
    End If
Next
AyWrpPad = O
End Function

Sub WrtAy(A, Ft, Optional OvrWrt As Boolean)
WrtStr JnCrLf(A), Ft, OvrWrt
End Sub

Function AyZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O()
    Dim J&
    ReszAyU O, U
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

Function AyZip_Ap(A, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
Dim UCol%
    UCol = UB(Av)

Dim URow1&
    URow1 = UB(A)

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
    ReszAyU ODry, URow
    Dim I%
    For J = 0 To URow
        Erase Dr
        If URow1 >= J Then
            Push Dr, A(J)
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


Function AyItmAddAy(Itm, Ay)
AyItmAddAy = AyInsItm(Ay, Itm)
End Function

Function SubDrFnySel(Dr(), DrFny$(), SelFF) As Variant()
Dim SelIxAy&()

SubDrFnySel = AywIxAy(Dr, SelIxAy)
End Function

Private Sub ZZZ_AyABCzAyFT()
Dim A(): A = Array(1, 2, 3, 4)
Dim Act As AyABC: Act = AyABCzAyFT(A, 1, 2)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Private Sub ZZ_AyAsgAp()
Dim O%, A$
'AyAsgAp Array(234, "abc"), O, A
Ass O = 234
Ass A = "abc"
End Sub

Private Sub ZZ_ChkEqAy()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub ZZ_MaxAy()
Dim A()
Dim Act
Act = MaxAy(A)
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
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyeEmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub ZZ_SyzAy()
Dim Act$(): Act = SyzAy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub ZZ_AyTrim()
DmpAy AyTrim(Array(1, 2, 3, "  a"))
End Sub


Private Sub Z_ChkAyDup()
Dim Ay
Ay = Array("1", "1", "2")
Ept = Sy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = ChkAyDup(Ay, "This item[?] is duplicated")
    C
    Return
End Sub

Private Sub Z_ChkEqAy()
DmpAy ChkEqAy(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub Z_AyABCzAyFTIxIx()
Dim A(): A = Array(1, 2, 3, 4)
Dim M As FTIx: M = FTIx(1, 2)
Dim Act As AyABC: Act = AyABCzAyFTIx(A, M)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Private Sub Z_HasEleDupEle()
Ass HasEleDupEle(Array(1, 2, 3, 4)) = False
Ass HasEleDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Private Sub Z_AyInsItm()
Dim A, M, At&
'
A = Array(1, 2, 3)
M = "X"
Ept = Array("X", 1, 2, 3)
GoSub Tst
'
Exit Sub
Tst:
    Act = AyInsItm(A, M, At)
    C
Return
End Sub

Private Sub Z_AyInsAy()
Dim Act, Exp, A(), B(), At&
A = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAyAt(A, B, At)
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
Ass Sz(Act) = 3
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
'Ass Sz(Act) = 4
'Ass IsEqAy(Act(0), Array("A", 3, 1, 2, 3))
'Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
'Ass IsEqAy(Act(2), Array("C", 0))
'Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub
Private Sub Z_SubDrFnySel()
Dim DrFny$(), SelFF$, Dr()
DrFny = SySsl("A B C D E F")
Dr = Array(Empty, Empty, 1, Empty, 2)
SelFF = "C E"
Ept = Array(1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = SubDrFnySel(Dr, DrFny, SelFF)
    C
    Return
End Sub

Function CvAy(A) As Variant()
CvAy = A
End Function

Function CvAyITM(Itm_or_Ay) As Variant()
If IsArray(Itm_or_Ay) Then
    CvAyITM = Itm_or_Ay
Else
    CvAyITM = Array(Itm_or_Ay)
End If
End Function


Private Sub Z()
Z_ChkAyDup
Z_AyFlat
Z_AyABCzAyFTIxIx
Z_HasEleDupEle
Z_ChkEqAy
Z_AyMinus
Z_SyzAy
Z_AyTrim
Z_SubDrFnySel
MVb_Ay:
End Sub

Private Sub Z_SyPfxSsl()
Dim A$, Exp$()
A = "A B C D"
Exp = SySsl("AB AC AD")
GoSub Tst
Exit Sub
Tst:
    Dim Act$()
    Act = SyPfxSsl(A)
    Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Function SyPfxSsl(A) As String()
Dim Ay$(), Pfx$
Ay = SySsl(A)
Pfx = AyShf(Ay)
SyPfxSsl = AyAddPfx(Ay, Pfx)
End Function

Function SySsl(S) As String()
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

