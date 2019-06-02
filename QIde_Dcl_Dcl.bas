Attribute VB_Name = "QIde_Dcl_Dcl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Dcl_Lines."
Private Const Asm$ = "QIde"
Public Const DoczDclDic$ = "Key is Pjn.Mdn.  Value is Dcl (which is Lines)"
Public Const DoczDcl$ = "It is Lines."

Function EnmBdyLyzS(Src$(), Enmn) As String()
EnmBdyLyzS = EnmBdyLy(EnmLy(Src, Enmn))
End Function

Function EnmBdyLy(EnmLy$()) As String()

End Function

Function EnmFei(Src$(), Enmn) As Fei
Dim Fm&: Fm = EnmFmIx(Src, Enmn)
EnmFei = Fei(Fm, EndLix(Src, Fm))
End Function

Function EnmLy(Src$(), Enmn) As String()
EnmLy = AywFei(Src, EnmFei(Src, Enmn))
End Function

Function EnmFmIx&(Src$(), Enmn)
Dim J&, L, Lin$
For Each L In Itr(Src)
    Lin = RmvMdy(L)
    If ShfTermEnm(Lin) Then
        If Nm(Lin) = Enmn Then
            EnmFmIx = J
            Exit Function
        End If
    End If
    If IsMthLin(Lin) Then Exit For
    J = J + 1
Next
EnmFmIx = -1
End Function

Function EnmNyMd(A As CodeModule) As String()
EnmNyMd = EnmNy(DclLyzM(A))
End Function
Function EnmNyPj(Pj As VBProject) As String()
Dim M
For Each M In MdItr(Pj)
    PushIAy EnmNyPj, EnmNyMd(CvMd(M))
Next
End Function
Function EnmNy(Src$()) As String()
Dim L
For Each L In Itr(Src)
   PushNonBlank EnmNy, Enmn(L)
Next
End Function

Function HasTyn(Src$(), Nm$) As Boolean
Dim L
For Each L In Itr(Src)
    If Tyn(L) = Nm Then HasTyn = True: Exit Function
Next
End Function

Function NEnm%(Src$())
Dim L, O%
For Each L In Itr(Src)
   If IsEmnLin(L) Then O = O + 1
Next
NEnm = O
End Function

Function UsrTyFei(Dcl$(), TyNm$) As Fei
Dim FmI&: FmI = UsrTyFmIx(Dcl, TyNm)
Dim ToI&: ToI = EndLix(Dcl, FmI)
UsrTyFei = Fei(FmI, ToI)
End Function

Function MthELin$(MthLin)
Dim K$: K = MthKd(MthLin)
If K = "" Then Thw CSub, "MthLin Error", "MthLin", MthLin
MthELin = "End " & K
End Function

Function UsrTyLines$(Dcl$(), Tyn$)
UsrTyLines = JnCrLf(UsrTyLy(Dcl, Tyn))
End Function

Function UsrTyLy(Dcl$(), TyNm$) As String()
UsrTyLy = AywFei(Dcl, UsrTyFei(Dcl, TyNm))
End Function

Function UsrTyFmIx&(Src$(), TyNm)
Dim J%
For J = 0 To UB(Src)
   If IsUsrTyLin(Src(J)) = TyNm Then UsrTyFmIx = J: Exit Function
   If IsMthLin(Src(J)) Then Exit For
Next
UsrTyFmIx = -1
End Function

Function TynyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlank TynyzS, TynzLin(L)
    If IsMthLin(L) Then Exit Function
Next
End Function

Function IsEmnLin(A) As Boolean
IsEmnLin = HasPfx(RmvMdy(A), "Enum ")
End Function

Function IsUsrTyLin(A) As Boolean
IsUsrTyLin = HasPfx(RmvMdy(A), "Type ")
End Function

Function Enmn(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Enum ") Then Enmn = Nm(LTrim(L))
End Function

Function Tyn$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfTermTy(L) Then Tyn = Nm(L)
End Function

Function EnmLyzMN(M As CodeModule, Enmn) As String()
EnmLyzMN = EnmLy(DclLyzM(M), Enmn)
End Function

Function NEnmMbrzMN%(M As CodeModule, Enmn)
NEnmMbrzMN = Si(EnmMbrLyzMN(M, Enmn))
End Function

Function CdLyzL(Lin) As String()
Dim L$: L = Trim(Lin)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
CdLyzL = TrimAy(Split(Lin, ":"))
End Function
Private Sub ZZ_CdLyzS()
Brw CdLyzS(SrczP(CPj))
End Sub
Function CdLyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushIAy CdLyzS, CdLyzL(L)
Next
End Function

Function EnmMbrLyzMN(M As CodeModule, Enmn) As String()
EnmMbrLyzMN = CdLyzS(EnmLyzMN(M, Enmn))
End Function

Function NEnmzM%(M As CodeModule)
NEnmzM = NEnm(DclLyzM(M))
End Function

Function TynyzM(M As CodeModule) As String()
TynyzM = TynyzS(DclLyzM(M))
End Function

Function TynyzP(P As VBProject) As String()
Dim I, C As VBComponent
For Each C In P.VBComponents
    PushIAy TynyzP, TynyzM(C.CodeModule)
Next
End Function

Function ShfTermEnm(OLin$) As Boolean
ShfTermEnm = ShfPfx(OLin, "Enum")
End Function

Function ShfTermTy(OLin$) As Boolean
ShfTermTy = ShfPfx(OLin, "Type")
End Function

Private Sub ZZ()
MIde_Dcl_EnmAndTy:
End Sub

Private Sub Z_NEnmMbrzMN()
Ass NEnmMbrzMN(Md("Ide"), "AA") = 1
End Sub

Private Sub Z_DclLinCnt()
Dim B1$(): B1 = CSrc
Dim B2$(): B2 = SrtSrc(B1)
Dim A1%: A1 = DclLinCnt(B1)
Dim A2%: A2 = DclLinCnt(B2)
Debug.Assert A1 = A2
End Sub

Sub BrwDclLinCntP()
BrwDry DclLinCntzP(CPj)
End Sub

Function DclLinCntzP(P As VBProject) As Variant()
Dim C As VBComponent
For Each C In P.VBComponents
    PushI DclLinCntzP, Array(C.Name, DclLinCntzM(C.CodeModule))
Next
End Function

Function DclLinCntzM%(M As CodeModule) 'Assume FstMth cannot have TopRmk
Dim I&
    I = FstMthLnozM(M)
    If I <= 0 Then
        DclLinCntzM = M.CountOfLines
        Exit Function
    End If
DclLinCntzM = TopRmkLno(M, I) - 1
End Function

Function DclLinCnt%(Src$()) 'Assume FstMth cannot have TopRmk
Dim Top&
    Dim Fm&
    Fm = FstMthIxzS(Src)
    If Fm = -1 Then
        DclLinCnt = UB(Src) + 1
        Exit Function
    End If
DclLinCnt = IxOfPrvCdLin(Src, Fm) + 1
End Function
Function IxOfPrvCdLin&(Src$(), Fm)
Dim O&
For O = Fm - 1 To 0 Step -1
    If IsCdLin(Src(O)) Then IxOfPrvCdLin = O: Exit Function
Next
IxOfPrvCdLin = -1
End Function
Function Dcl$(Src$())
Dcl = JnCrLf(DclLy(Src))
End Function

Function DclDicP() As Dictionary
Set DclDicP = DclDiczP(CPj)
End Function

Function DclDiczP(P As VBProject) As Dictionary
If P.Protection = vbext_pp_locked Then Set DclDiczP = New Dictionary: Exit Function
Dim C As VBComponent, M As CodeModule
Set DclDiczP = New Dictionary
For Each C In P.VBComponents
    Set M = C.CodeModule
    Dim Dcl$: Dcl = DclzM(M)
    If Dcl <> "" Then
        DclDiczP.Add MdDNm(M), Dcl
    End If
Next
End Function

Function DclItr(A As CodeModule)
Asg Itr(DclLyzM(A)), DclItr
End Function

Function DclLy(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim N&, O$()
   N = DclLinCnt(Src)
If N <= 0 Then Exit Function
O = AywFstNEle(Src, N)
DclLy = O
'Brw LyzNNAp("N Src DclLy", N, AddIxPfx(Src), O): Stop
End Function
Function LineszMLC$(M As CodeModule, Lno&, Cnt&)
If Lno <= 0 Then Exit Function
If Cnt <= 0 Then Exit Function
If Lno > M.CountOfLines Then Exit Function
LineszMLC = M.Lines(Lno, Cnt)
End Function
Private Sub ZZ_DclzM()
Dim O$(), C As VBComponent
For Each C In CPj.VBComponents
    PushNonBlank O, DclzM(C.CodeModule)
Next
VcLinesAy O
End Sub
Function DclzM$(A As CodeModule)
DclzM = TrimRSpcCrLf(LineszMLC(A, 1, DclLinCntzM(A)))
End Function
Function DclLyzM(A As CodeModule) As String()
DclLyzM = SplitCrLf(DclzM(A))
End Function
Function CnstLnxszS(Src$()) As Lnxs
Dim L, J&
For Each L In Itr(Src)
    If IsLin_OfCnst(L) Then PushLnx CnstLnxszS, Lnx(L, J)
    J = J + 1
Next
End Function

Function CnstLnxzSN(Src$(), CnstnPfx$) As Lnx
Dim L, J%
For Each L In Itr(Src)
    If IsLin_OfCnst_WhNmPfx(L, CnstnPfx) Then CnstLnxzSN = Lnx(L, J): Exit Function
    J = J + 1
Next
End Function
Function IsLin_OfCnst_WhNmPfx(L, CnstnPfx$) As Boolean
Dim Lin$: Lin = RmvMdy(L)
If Not ShfTermCnst(Lin) Then Exit Function
IsLin_OfCnst_WhNmPfx = HasPfx(L, CnstnPfx)
End Function
Function IsLin_OfCnst(L) As Boolean
IsLin_OfCnst = T1(RmvMdy(L)) = "Const"
End Function
Function CnstLnxszM(M As CodeModule) As Lnxs
Dim J&, L$, P$, L1$, L2$
P = "Const "
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    L1 = RmvMdy(L)
    If HasPfx(L1, P) Then
        L2 = ContLinzML(M, J)
        PushLnx CnstLnxszM, Lnx(L, J - 1)
    End If
Next
End Function

Function CnstLnxzMN(M As CodeModule, Cnstn$) As Lnx
Dim J&, L$, P$, L1$, L2$
P = "Const " & Cnstn
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    L1 = RmvMdy(L)
    If HasPfx(L1, P) Then
        L2 = ContLinzML(M, J)
        CnstLnxzMN = Lnx(L2, J - 1)
        Exit Function
    End If
Next
End Function


