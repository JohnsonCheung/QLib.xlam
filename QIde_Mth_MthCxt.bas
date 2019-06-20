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

Function DMthCxt() As Drs
DMthCxt = DMthCxtzML(CMd, CMthLno)
End Function

Function DMthnM() As Drs
DMthnM = DMthn(CMd)
End Function

Function DMthn(M As CodeModule) As Drs
DMthn = DrpCol(DMth(M), "MthLin")
End Function
Function AddColzHasPm(A As Drs) As Drs
'Fm A : ..MthLin..
'Ret  : ..MthLin.. HasPm
Dim I%: I = IxzAy(A.Fny, "MthLin")
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dim MthLin$: MthLin = Dr(I)
    Dim HasPm As Boolean: HasPm = BetBkt(MthLin) <> ""
    PushI Dr, HasPm
    PushI Dry, Dr
Next
AddColzHasPm = DrszAddFF(A, "HasPm", Dry)
End Function

Sub Z_TblMthP()
Dim D As Database: Set D = TmpDb
Dim T$: T = TblMthP(D)
BrwDb D
End Sub

Function TblMthP$(D As Database)
Dim T$: T = "Mth"
CrtTzDrs D, T, DMthP
TblMthP = T
End Function

Function DMthP() As Drs
Static A As Drs
If NoReczDrs(A) Then A = DMthzP(CPj)
DMthP = A
End Function

Function DMthzP(P As VBProject) As Drs
Dim C As VBComponent, ODry(), Dry(), Pjn$
Pjn = P.Name
For Each C In P.VBComponents
    Dry = DMth(C.CodeModule).Dry
    Dry = InsColzDryAv(Dry, Av(Pjn, ShtCmpTy(C.Type), C.Name))
    PushIAy ODry, Dry
Next
DMthzP = DrszFF("Pjn MdTy Mdn L Mdy Ty Mthn MthLin", ODry)
End Function

Function DMth(M As CodeModule) As Drs
'Ret : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht @@
DMth = DMthzS(Src(M))
End Function
Function DMthzM(M As CodeModule) As Drs
'Ret : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht @@
DMthzM = DMthzS(Src(M))
End Function

Function DMthe(M As CodeModule) As Drs
'Ret : L E Mdy Ty Mthn MthLin ! Mdy & Ty are Sht. L is Lno E is ELno @@d
DMthe = DMthezS(Src(M))
End Function

Function DMthc(M As CodeModule) As Drs
'Ret : L E Mdy Ty Mthn MthLin MthLy! Mdy & Ty are Sht. L is Lno E is ELno @@
DMthc = DMthczS(Src(M))
End Function

Function DMtheM() As Drs
DMtheM = DMthe(CMd)
End Function
Function DMthcM() As Drs
DMthcM = DMthczM(CMd)
End Function

Function DMthczS(Src$()) As Drs
Dim A As Drs: A = DMthzS(Src)
Dim Dry(), Dr
For Each Dr In Itr(A.Dry)
    Dim F&: F = Dr(0) - 1
    Dim E&: E = EndLix(Src, F) + 1
    Dim T&: T = E - 1
    Dim MthLy$(): MthLy = AywFT(Src, F, T)
    Dr = InsEle(Dr, E, 1)
    PushI Dr, MthLy
    PushI Dry, Dr
Next
DMthczS = DrszFF("L E Mdy Ty Mthn MthLin MthLy", Dry)
End Function

Function AddColzRetAs(DMthLin As Drs) As Drs
'Fm DMthLin : ..MthLin..
'Ret        : ..RetAs
Dim IxMthLin%: IxMthLin = IxzAy(DMthLin.Fny, "MthLin")
Dim Dr, Dry(): For Each Dr In Itr(DMthLin.Dry)
    Dim MthLin$: MthLin = Dr(IxMthLin)
    Dim R$: R = RetAszL(MthLin)
    PushI Dr, R
    PushI Dry, Dr
Next
AddColzRetAs = DrszAddFF(DMthLin, "RetAs", Dry)
End Function
Function DMthezS(Src$()) As Drs
Dim A As Drs: A = DMthzS(Src)
Dim Dry(), Dr
For Each Dr In Itr(A.Dry)
    Dim Ix&: Ix = Dr(0) - 1
    Dim E&: E = EndLix(Src, Ix) + 1
    Dr = InsEle(Dr, E, 1)
    PushI Dry, Dr
Next
DMthezS = DrszFF("L E Mdy Ty Mthn MthLin", Dry)
End Function

Function DMthzS(Src$()) As Drs
'Ret DMth : L Mdy Ty Mthn MthLin ! Mdy & Ty are Sht
Dim Dry(), Dr, Ix
For Each Ix In MthIxItr(Src)
    Dim L&:              L = Ix + 1
    Dim MthLin$:    MthLin = ContLin(Src, Ix)
    Dim A As Mthn3:      A = Mthn3zL(MthLin)
    Dim Ty$:            Ty = A.ShtTy
    Dim Mdy$:          Mdy = A.ShtMdy
    Dim Mthn$:        Mthn = A.Nm
    PushI Dry, Array(L, Mdy, Ty, Mthn, MthLin)
Next
DMthzS = DrszFF("L Mdy Ty Mthn MthLin", Dry)
End Function

Function DMthCxtzML(M As CodeModule, MthLno&) As Drs
'Ret DMthCxt : L Lin
Dim Dry(), L&, ELin$, MthLin$, Lin$
MthLin = M.Lines(MthLno, 1)
ELin = MthELin(MthLin)
For L = NxtLnozML(M, MthLno) To M.CountOfLines
    Lin = M.Lines(L, 1)
    If Lin = ELin Then
        GoTo X
    End If
    Lin = M.Lines(L, 1)
    PushI Dry, Array(L, Lin)
Next
ThwImpossible CSub
X:
DMthCxtzML = DrszFF("L MthLin", Dry)
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

Sub Z_InspMthStmt()
Dim A As Drs: A = DMthzM(CMd)
Dim B$(): B = StrCol(A, "MthLin")
Dim L, ODry()
For Each L In B
    PushI ODry, Array(L, InspMthStmt(L, "Md"))
Next
Dim C As Drs: C = DrszFF("MthLin InspStmt", ODry)
Brw LinzDrsR(C)
End Sub

Function InspMthStmt$(MthLin, Mdn$)
With MthLinRec(MthLin)
    If .Pm = "" And Not .IsRetVal Then Exit Function
    Dim NN$: NN = JnSpc(ArgNyzPm(.Pm))
    Dim EE$: EE = InspExprLiszPm(.Pm)
    Dim IsN0$: IsN0 = XIsN0(.IsRetVal, .Nm)  '#Insp-Nm-0.
    Dim IsE0$: IsE0 = XIsE0(.IsRetVal, .Nm, .TyChr, .RetTy) '#Insp-Expr-0
    NN = IsN0 & NN
    EE = IsE0 & EE
    InspMthStmt = InspStmt(NN, EE, Mdn, .Nm)
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
Case "DLTDH": O = FmtQQ("LinzDrs(?.D)", V)
Case "Drs": O = FmtQQ("LinzDrs(?)", V)
Case "S1S2s": O = FmtQQ("FmtS1S2s(?)", V)
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

