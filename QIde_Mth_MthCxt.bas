Attribute VB_Name = "QIde_Mth_MthCxt"
Option Compare Text
Option Explicit
Type Fc
    FmLno As Long
    Cnt As Long
End Type
Type Fcs: N As Long: Ay() As Fc: End Type
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Rmk."

Sub UnRmkMth(M As CodeModule, Mthn)
'UnRmkMdzFes A, MthCxtFes(Src(A), Mthn)
End Sub

Function Fc(FmLno, Cnt) As Fc
If Cnt <= 0 Then Exit Function
If FmLno <= 0 Then Exit Function
Fc.FmLno = FmLno
Fc.Cnt = Cnt
End Function

Sub RmkLineszFc(M As CodeModule, Fc As Fc)
RmkLines M, Fc.FmLno, Fc.Cnt
End Sub

Function MthCxtFcs(Src$(), Mthn) As Fcs
MthCxtFcs = MthCxtFcszzSM(Src, MthCxtFcs(Src, Mthn))
End Function

Function MthCxtFcszzSM(Src$(), Mth As Fcs) As Fcs
Dim J&
For J = 0 To Mth.N - 1
    PushFc MthCxtFcszzSM, MthCxtFczzSM(Src, Mth.Ay(J))
Next
End Function

Function NContLin(Src$(), MthIx) As Byte
Dim J&, O&
For J = MthIx To UB(Src)
    O = O + 1
    If LasChr(Src(J)) <> "_" Then NContLin = O: Exit Function
Next
Thw CSub, "LasEle of Src has LasChr = _", "Src", Src
End Function

Function AddFc(A As Fc, B As Fc) As Fcs
PushFc AddFc, A
PushFc AddFc, B
End Function

Function FmtFcs$(A As Fcs)
Dim O$(), J&
For J = 0 To A.N - 1
    PushI O, FmtFc(A.Ay(J))
Next
FmtFcs = JnCrLf(O)
End Function

Function FmtFc$(Fc As Fc)
With Fc
FmtFc = "Fc " & .FmLno & " " & .Cnt
End With
End Function

Function MthCxtFczzSM(Src$(), Mth As Fc) As Fc
With Mth
Dim N%: N = NContLin(Src, .FmLno)
MthCxtFczzSM = Fc(.FmLno - N, .Cnt - N - 1)
End With
End Function
Function MthCxtFc(M As CodeModule, Mthn) As Fc

End Function
Sub RmkMthzN(M As CodeModule, Mthn)
RmkLineszFc M, MthCxtFc(M, Mthn)
End Sub

Function MthFcs(M As CodeModule, Mthn) As Fcs
Dim Ix, S$()
S = Src(M)
For Each Ix In Itr(MthIxyzSN(S, Mthn))
    PushFc MthFcs, Fc(Ix + 1, ContLinCnt(S, Ix))
Next
End Function

Sub PushFc(O As Fcs, M As Fc)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Sub RmkLines(M As CodeModule, Lno&, N&)
Dim J&
For J = Lno To Lno + N - 1
    RmkLin M, J
Next
End Sub

Sub RmkLin(M As CodeModule, Lno&)
M.ReplaceLine M, "'" & M.Lines(Lno, 1)
End Sub

Sub RmkMth()
RmkMthzN CMd, CMthn
End Sub

Private Sub Z_RmkMth()
Dim Md As CodeModule, Mthn
'            Ass LineszVbl(MthL(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
'RmkMth M:   Ass LineszVbl(MthL(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
'UnRmkMth M: Ass LineszVbl(MthL(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
Function NxtMdLno(M As CodeModule, Lno)
Const CSub$ = CMod & "NxtMdLno"
Dim J&
For J = Lno To M.CountOfLines
    If LasChr(M.Lines(Lno, 1)) <> "_" Then
        NxtMdLno = J
        Exit Function
    End If
Next
Thw CSub, "All line From Lno has _ as LasChr", "Lno Md Src", Lno, Mdn(M), AddIxPfx(Src(M), 1)
End Function

Sub UnRmkMdzFes(M As CodeModule, B As Feis)
Dim J&
For J = 0 To B.N - 1
    UnRmkMdzFei M, B.Ay(J)
Next
End Sub

Sub UnRmkMdzFei(M As CodeModule, B As Fei)
'If Not IsRmkedzS(LyzMdFei(A, B)) Then Exit Sub
Stop
Dim J%, L$
'For J = NxtMdLno(A, B.FmNo) To B.ToNo - 1
    L = M.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    M.ReplaceLine J, Mid(L, 2)
'Next
End Sub

Sub RmkMdzFes(M As CodeModule, B As Feis)
Dim J%
For J = 0 To B.N - 1
    RmkMdzFe M, B.Ay(J)
Next
End Sub

Sub RmkMdzFe(M As CodeModule, B As Fei)
If IsRmkedzMFe(M, B) Then Exit Sub
Dim J%
'For J = 0 To UB(B)
    M.ReplaceLine J, "'" & M.Lines(J, 1)
'Next
End Sub

Function IsRmkedzMFe(M As CodeModule, B As Fei) As Boolean
'IsRmkedzMFe = IsRmkedzS(LyzMdFei(A, B))
End Function

Function DoMthCxt() As Drs
DoMthCxt = DoMthCxtzML(CMd, CMthLno)
End Function

Function DoMthnM() As Drs
DoMthnM = DoMthn(CMd)
End Function

Function DoMthn(M As CodeModule) As Drs
DoMthn = DrpCol(DoMthzM(M), "MthLin")
End Function

Function AddColzHasPm(A As Drs) As Drs
'Fm A : ..MthLin..
'Ret  : ..MthLin.. HasPm
Dim I%: I = IxzAy(A.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim HasPm As Boolean: HasPm = BetBkt(MthLin) <> ""
    PushI Dr, HasPm
    PushI Dy, Dr
Next
AddColzHasPm = AddColzFFDy(A, "HasPm", Dy)
End Function

Private Sub Z_TblMthP()
Dim D As Database: Set D = TmpDb
Dim T$: T = TblMthP(D)
BrwDb D
End Sub

Function TblMthP$(D As Database)
Dim T$: T = "Mth"
CrtTzDrs D, T, DoPubMth
TblMthP = T
End Function

Function AddColzBetBkt(D As Drs, ColnAs$, Optional IsDrp As Boolean) As Drs
Dim BetColn$, NewC$: AsgBrk1 ColnAs, ":", BetColn, NewC
If NewC = "" Then NewC = BetColn & "InsideBkt"
Dim Ix%: Ix = IxzAy(D.Fny, BetColn)
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    PushI Dr, BetBkt(Dr(Ix))
    PushI Dy, Dr
Next
Dim O As Drs: O = AddColzFFDy(D, NewC, Dy)
If IsDrp Then O = DrpCol(O, BetColn)
AddColzBetBkt = O
End Function

Function AddColzFstArgSfx(D As Drs) As Drs
Dim Ix%: Ix = IxzAy(D.Fny, "MthPm")
Dim Dr, Dy(): For Each Dr In Itr(D.Dy)
    PushI Dr, ArgSfx(Bef(Dr(Ix), ","))
    PushI Dy, Dr
Next
Dim O As Drs: O = AddColzFFDy(D, "FstArgSfx", Dy)
AddColzFstArgSfx = O
End Function

Function AddColzMthPm(DoMthLin As Drs, Optional IsDrp As Boolean) As Drs
AddColzMthPm = AddColzBetBkt(DoMthLin, "MthLin:MthPm", IsDrp)
End Function

Function AddColzRetAs(DoMthLin As Drs) As Drs
'Fm DoMthLin : ..MthLin..
'Ret        : ..RetAs
Dim IxMthLin%: IxMthLin = IxzAy(DoMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(DoMthLin.Dy)
    Dim MthLin$: MthLin = Dr(IxMthLin)
    Dim R$: R = RetAszL(MthLin)
    PushI Dr, R
    PushI Dy, Dr
Next
AddColzRetAs = AddColzFFDy(DoMthLin, "RetAs", Dy)
End Function
Function DoMthezS(Src$()) As Drs
Dim A As Drs: A = DoMthzS(Src)
Dim Dy(), Dr
For Each Dr In Itr(A.Dy)
    Dim Ix&: Ix = Dr(0) - 1
    Dim E&: E = EndLix(Src, Ix) + 1
    Dr = InsEle(Dr, E, 1)
    PushI Dy, Dr
Next
DoMthezS = DrszFF("L E Mdy Ty Mthn MthLin", Dy)
End Function

Function DoMthzS(Src$()) As Drs
'Ret DoMth : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht
Dim Dy(), Dr, Ix
For Each Ix In MthIxItr(Src)
    Dim L&:              L = Ix + 1
    Dim MthLin$:    MthLin = ContLin(Src, Ix)
    Dim A As Mthn3:      A = Mthn3zL(MthLin)
    Dim Ty$:            Ty = A.ShtTy
    Dim Mdy$:          Mdy = A.ShtMdy
    Dim Mthn$:        Mthn = A.Nm
    PushI Dy, Array(L, Mdy, Ty, Mthn, MthLin)
Next
DoMthzS = DrszFF("L Mdy Ty Mthn MthLin", Dy)
End Function

Function DoMthCxtzML(M As CodeModule, MthLno&) As Drs
'Ret DoMthCxt : L Lin
Dim Dy(), L&, ELin$, MthLin$, Lin$
MthLin = M.Lines(MthLno, 1)
ELin = MthELin(MthLin)
For L = NxtLnozML(M, MthLno) To M.CountOfLines
    Lin = M.Lines(L, 1)
    If Lin = ELin Then
        GoTo X
    End If
    Lin = M.Lines(L, 1)
    PushI Dy, Array(L, Lin)
Next
ThwImpossible CSub
X:
DoMthCxtzML = DrszFF("L MthLin", Dy)
End Function
Function IsRmkedzMthLy(MthLy$()) As Boolean
If Si(MthLy) = 0 Then Exit Function
If Not HasPfx(MthLy(0), "Stop '") Then Exit Function
Dim L
For Each L In MthLy
    If Left(L, 1) <> "'" Then Exit Function
Next
IsRmkedzMthLy = True
End Function
Function MthCxtFe(MthLy$(), Fe As Fei) As Fei
MthCxtFe = Fei(NxtIxzSrc(MthLy, Fe.FmIx), Fe.EIx - 1)
End Function

Function MthCxt$(MthLy$())
MthCxt = JnCrLf(MthCxtLy(MthLy))
End Function

Function MthCxtLy(MthLy$()) As String()
If Si(MthLy) = 0 Then Exit Function
Dim L&: L = FstMthIx(MthLy): If L = -1 Then Thw CSub, "Given MthLy is not MthLy", "MthLy", MthLy
Dim J%
For J = NxtIxzSrc(MthLy, L) To UB(MthLy) - 1
    PushI MthCxtLy, MthLy(J)
Next
End Function

Private Sub Z_MthCxtFeis _
 _
()
Stop
Dim I
'For Each I In MthCxtFeis(CSrc, CurMthn)
    'With CvFei(I)
'        Debug.Print .FmNo, .ToNo
    'End With
'Next
End Sub

Function InspExprLiszPm$(Pm$)
If Pm = "" Then Exit Function
Dim Ay$(): Ay = Split(Pm, ", ")
Dim O$(), P
For Each P In Ay
    Dim L$: L = P
    Dim V$: V = ShfNm(L)
    Dim S$: S = L
    PushI O, InspExprzDclSfx(V, S)
Next
InspExprLiszPm = JnCommaSpc(O)
End Function

Private Sub Z_InspMthStmt()
Dim A As Drs: A = DoMthzM(CMd)
Dim B$(): B = StrCol(A, "MthLin")
Dim L, ODy()
For Each L In B
    PushI ODy, Array(L, InspMthStmt(L, "Md"))
Next
Dim C As Drs: C = DrszFF("MthLin InspStmt", ODy)
Brw LinzDrsR(C)
End Sub

Function InspMthStmt$(MthLin, Mdn$)
With MthLinRec(MthLin)
    If .Pm = "" And Not .IsRetVal Then Exit Function
    Dim NN$: NN = JnSpc(ArgNyzPm(.Pm))
    Dim Ee$: Ee = InspExprLiszPm(.Pm)
    Dim IsN0$: IsN0 = XIsN0(.IsRetVal, .Nm)  '#Insp-Nm-0.
    Dim IsE0$: IsE0 = XIsE0(.IsRetVal, .Nm, .TyChr, .RetTy) '#Insp-Expr-0
    NN = IsN0 & NN
    Ee = IsE0 & Ee
    InspMthStmt = InspStmt(NN, Ee, Mdn, .Nm)
End With
End Function

Private Function XIsN0$(IsRetVal As Boolean, Mthn$)
If Not IsRetVal Then Exit Function
XIsN0 = "Oup(" & Mthn & ") "
End Function

Private Function XIsE0$(IsRetVal As Boolean, V, TyChr$, RetTy$)
If Not IsRetVal Then Exit Function
XIsE0 = InspExprzDclSfx(V, TyChr & RetTy) & ", "
End Function

Function InspStmt$(NN$, ExprLis$, Mdn$, Mthn$)
Const C$ = "Insp ""?.?"", ""Inspect"", ""?"", ?"
InspStmt = FmtQQ(C, Mdn, Mthn, NN, ExprLis)
End Function

Private Function InspExpr$(V, VSfx As Dictionary)
If Not VSfx.Exists(V) Then
    InspExpr = FmtQQ("""V(?)-NFnd""", V)
    Exit Function
End If
InspExpr = InspExprzDclSfx(V, VSfx(V))
End Function
Private Function InspExprzDclSfx$(V, DclSfx$)
Dim O$, S$
S = RmvPfx(DclSfx, " As ")
Select Case S
Case "DoLTDH": O = FmtQQ("FmtDrs(?.D)", V)
Case "Drs": O = FmtQQ("FmtDrs(?)", V)
Case "S12s": O = FmtQQ("FmtS12s(?)", V)
Case "CodeModule": O = FmtQQ("Mdn(?)", V)
Case "", "$", "$()", "#", "@", "%", "&", "%()", "&()", "#()", "@()", "$()": O = V
Case "Dictionary", "Byte", "Boolean", "String", "Integer": O = V
Case "", "String()", "Integer()", "Long()", "Byte()": O = V
Case Else: O = """NoFmtr(" & S & ")"""
End Select
InspExprzDclSfx = O
Exit Function
X: InspExprzDclSfx = FmtQQ(Q, V)
End Function
Function InspExprLis$(PP$, VSfx As Dictionary)
InspExprLis = Join(InspExprs(PP, VSfx), ", ")
End Function

Private Function InspExprs(PP$, VSfx As Dictionary) As String()
Dim V
For Each V In Itr(SyzSS(PP))
    PushI InspExprs, InspExpr(V, VSfx)
Next
End Function

