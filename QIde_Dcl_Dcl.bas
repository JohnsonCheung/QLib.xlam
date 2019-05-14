Attribute VB_Name = "QIde_Dcl_Dcl"
Option Explicit
Private Const CMod$ = "MIde_Dcl_Lines."
Private Const Asm$ = "QIde"
Public Const DoczDclDic$ = "Key is Pjn.Mdn.  Value is Dcl (which is Lines)"
Public Const DoczDcl$ = "It is Lines."

Function EnmBdyLyzSrc(Src$(), EnmNm$) As String()
EnmBdyLyzSrc = EnmBdyLy(EnmLy(Src, EnmNm$))
End Function

Function EnmBdyLy(EnmLy$()) As String()

End Function

Function EnmFEIx(Src$(), EnmNm) As FEIx
Dim Fm&: Fm = EnmFmIx(Src, EnmNm)
EnmFEIx = FEIx(Fm, EndEnmIx(Src, Fm))
End Function

Function EnmLy(Src$(), EnmNm$) As String()
EnmLy = AywFEIx(Src, EnmFEIx(Src, EnmNm))
End Function

Function EnmFmIx&(Src$(), EnmNm)
Dim J&, L, Lin$
For Each L In Itr(Src)
    Lin = RmvMdy(L)
    If ShfTermEnm(Lin) Then
        If Nm(Lin) = EnmNm Then
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
EnmNyMd = EnmNy(DclLyzMd(A))
End Function
Function EnmNyPj(Pj As VBProject, Optional WhStr$) As String()
Dim M
For Each M In MdItr(Pj, WhStr)
    PushIAy EnmNyPj, EnmNyMd(CvMd(M))
Next
End Function
Function EnmNy(Src$()) As String()
Dim L
For Each L In Itr(Src)
   PushNonBlank EnmNy, EnmNm(L)
Next
End Function

Function HasUsrTyNm(Src$(), Nm$) As Boolean
Dim L
For Each L In Itr(Src)
    If UsrTyNm(L) = Nm Then HasUsrTyNm = True: Exit Function
Next
End Function

Function NEnm%(Src$())
Dim L, O%
For Each L In Itr(Src)
   If IsEmnLin(L) Then O = O + 1
Next
NEnm = O
End Function

Function UsrTyFEIx(Src$(), TyNm$) As FEIx
Dim FmI&: FmI = UsrTyFmIx(Src, TyNm)
Dim ToI&: ToI = EndTyIx(Src, FmI)
UsrTyFEIx = FEIx(FmI, ToI)
End Function

Function EndEnmIx&(Src$(), FmIx)
EndEnmIx = EndLinIx(Src, "Enum", FmIx)
End Function

Function EndTyIx&(Src$(), FmIx)
EndTyIx = EndLinIx(Src, "Type", FmIx)
End Function

Function UsrTyLines$(Src$(), UsrTyNm$)
UsrTyLines = JnCrLf(UsrTyLy(Src, UsrTyNm))
End Function

Function UsrTyLy(Src$(), TyNm$) As String()
UsrTyLy = AywFEIx(Src, UsrTyFEIx(Src, TyNm))
End Function

Function UsrTyFmIx&(Src$(), TyNm)
Dim J%
For J = 0 To UB(Src)
   If IsUsrTyLin(Src(J)) = TyNm Then UsrTyFmIx = J: Exit Function
   If IsMthLin(Src(J)) Then Exit For
Next
UsrTyFmIx = -1
End Function

Function TyNyzS(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlank TyNyzS, TynzLin(L)
    If IsMthLin(L) Then Exit Function
Next
End Function

Function IsEmnLin(A) As Boolean
IsEmnLin = HasPfx(RmvMdy(A), "Enum ")
End Function

Function IsUsrTyLin(A) As Boolean
IsUsrTyLin = HasPfx(RmvMdy(A), "Type ")
End Function

Function EnmNm$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Enum ") Then EnmNm = Nm(LTrim(L))
End Function

Function UsrTyNm$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Type ") Then UsrTyNm = Nm(LTrim(L))
End Function

Function EnmLyMd(Md As CodeModule, EnmNm$) As String()
EnmLyMd = EnmLy(DclLyzMd(Md), EnmNm)
End Function

Function NEnmMbrzMN%(A As CodeModule, Enmn$)
NEnmMbrzMN = Si(EnmMbrLyzMN(A, EnmNm))
End Function
Function CdLyzSrc(Src$()) As String()

End Function
Function EnmMbrLyzMN(A As CodeModule, Enmn$) As String()
EnmMbrLyzMN = CdLyzSrc(EnmLyzMN(A, Enmn))
End Function

Function NEnmzM%(A As CodeModule)
NEnmMd = NEnm(DclLyzMd(A))
End Function

Function TyNyzM(A As CodeModule) As String()
UsrTyNyMd = AySrt(UsrTyNy(DclLyzMd(A)))
End Function

Function TyNyzP(P As VBProject, Optional WhStr$) As String()
Dim I, M As CodeModule, O$(), W As WhNm
Set W = WhNmzStr(WhStr)
'For Each I In MdItr(A, WhStr)
    Set M = CvMd(I)
    O = UsrTyNy(Src(M))
    O = AywNm(O, W)
    PushIAy UsrTyNyPj, AddPfxzAy(O, Mdn(M) & ".")
'Next
UsrTyNyPj = QSrt1(O)
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

Private Sub Z_NEnmMbrMd()
Ass NEnmMbrMd(Md("Ide"), "AA") = 1
End Sub


Private Sub Z_DclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = SrcSrt(B1)
Dim A1%: A1 = DclLinCnt(B1)
Dim A2%: A2 = DclLinCnt(B2)
End Sub

Sub BrwDclLinCntDryPj()
BrwDry DclLinCntDryzP(CPj)
End Sub

Function DclLinCntDryzP(P As VBProject) As Variant()
Dim C As VBComponent
For Each C In P.VBComponents
    PushI DclLinCntDryzP, Array(C.Name, DclLinCntzMd(C.CodeModule))
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
Function DclLyzMd(A As CodeModule) As String()
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
IsLin_OfCnst_WhNmPfx = HasPfx(L, NmPfx)
End Function
Function IsLin_OfCnst(L) As Boolean
IsLin_OfCnst = T1(RmvMdy(L)) = "Const"
End Function
Function CnstLnxszM(M As CodeModule) As Lnxs
Dim J&, L$, P$, L1$, L2$
P = "Const " & Cnstn
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    L1 = RmvMdy(L)
    If HasPfx(L1, P) Then
        L2 = ContLinzML(M, J)
        PushLnx CnstLnxszM, Lnx(L2, J - 1)
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


