Attribute VB_Name = "QIde_B_DclItm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Dim."
Private Const Asm$ = "QIde"

Function DclNm$(DclItm$)
If HasSubStr(DclItm, " As ") Then
    DclNm = DclNm__As(DclItm)
Else
    DclNm = DclNm__TyChr(DclItm)
End If
End Function

Private Function DclNm__TyChr$(DimShtItm$)
DclNm__TyChr = RmvLasChrzzLis(RmvSfxzBkt(DimShtItm), MthTyChrLis)
End Function

Private Function DclNm__As$(DimAsItm$)
DclNm__As = RmvSfxzBkt(Bef(DimAsItm, " As"))
End Function

Function DclItmAyzDimLin(DimLin) As String()
Dim L$: L = DimLin
If Not ShfPfx(L, "Dim ") Then Exit Function
DclItmAyzDimLin = SplitCommaSpc(L)
End Function

Function DclNy(DclItmAy$()) As String()
Dim DclItm$, I
For Each I In Itr(DclItmAy)
    DclItm = I
    PushI DclNy, DclNm(DclItm)
Next
End Function

Function MdPosStr$(A As MdPos)
Dim B$
With A
    'With .LinPos.Pos
        'If .Cno1 > 0 Then B = " " & .Cno1 & " " & .Cno2
    'End With
    'MdPosStr = "MdPos " & Mdn(A.Md) & A.LinPos.Lno & B
End With
End Function

Function MdPoszMLCC(Md As CodeModule, L, Cno1, Cno2) As MdPos
'MdPoszMLCC = MdPos(Md, LinPoszLCC(L, Cno1, Cno2))
End Function

Function MdPoszMLP(Md As CodeModule, Lno, P As Pos) As MdPos
'MdPoszMLP = MdPos(Md, LinPos(Lno, P))
End Function

Function MdPos(Md As CodeModule, RRCC As RRCC) As MdPos
Set MdPos.Md = Md
MdPos.RRCC = RRCC
End Function


Function NTyMd%(M As CodeModule)
NTyMd = NTySrc(DclLyzM(M))
End Function


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
EnmLy = AwFei(Src, EnmFei(Src, Enmn))
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
    If IsLinMth(Lin) Then Exit For
    J = J + 1
Next
EnmFmIx = -1
End Function

Function EnmNyMd(M As CodeModule) As String()
EnmNyMd = EnmNy(DclLyzM(M))
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
   PushNB EnmNy, Enmn(L)
Next
End Function

Function HasTyn(Src$(), Nm$) As Boolean
Dim L
For Each L In Itr(Src)
    If Mtyn(L) = Nm Then HasTyn = True: Exit Function
Next
End Function

Function NEnm%(Src$())
Dim L, O%
For Each L In Itr(Src)
   If IsLinEmn(L) Then O = O + 1
Next
NEnm = O
End Function

Function TyFei(Dcl$(), TyNm$) As Fei
Dim FmI&: FmI = TyFmIx(Dcl, TyNm)
Dim ToI&: ToI = EndLix(Dcl, FmI)
TyFei = Fei(FmI, ToI)
End Function
Function MthELno&(M As CodeModule, MthLno&)
Dim MLin$: MLin = M.Lines(MthLno, 1)
Dim ELin$: ELin = MthELin(MLin)
Dim O&: For O = MthLno + 1 To M.CountOfLines
    If HasPfx(M.Lines(O, 1), ELin) Then MthELno = O: Exit Function
Next
End Function
Function MthELin$(MthLin)
Dim K$: K = MthKd(MthLin)
If K = "" Then Thw CSub, "MthLin Error", "MthLin", MthLin
MthELin = "End " & K
End Function

Function TyLines$(Dcl$(), Mtyn$)
TyLines = JnCrLf(TyLy(Dcl, Mtyn))
End Function

Function TyLy(Dcl$(), TyNm$) As String()
TyLy = AwFei(Dcl, TyFei(Dcl, TyNm))
End Function

Function TyFmIx&(Src$(), TyNm)
Dim J%
For J = 0 To UB(Src)
   If IsLinTy(Src(J)) = TyNm Then TyFmIx = J: Exit Function
   If IsLinMth(Src(J)) Then Exit For
Next
TyFmIx = -1
End Function

Function TyNy(DclLy$()) As String()
Dim L
For Each L In Itr(DclLy)
    PushNB TyNy, Mtyn(L)
Next
End Function
Function Enmn(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Enum ") Then Enmn = Nm(L)
End Function

Function Mtyn$(Lin)
':Mtyn: :Nm #Type-Name# ! Vb Type Name of @Lin
Dim L$: L = RmvMdy(Lin)
If ShfPfx(L, "Type ") Then Mtyn = Nm(L)
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
CdLyzL = SyzTrim(Split(Lin, ":"))
End Function
Private Sub Z_CdLyzS()
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

Function TyNyzM(M As CodeModule) As String()
TyNyzM = TyNy(DclLyzM(M))
End Function

Function TyNyP() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = TyNyzP(CPj)
End If
TyNyP = Y
End Function

Function TyNyzP(P As VBProject) As String()
Dim I, C As VBComponent
For Each C In P.VBComponents
    PushIAy TyNyzP, TyNyzM(C.CodeModule)
Next
End Function

Function ShfTermEnm(OLin$) As Boolean
ShfTermEnm = ShfPfx(OLin, "Enum")
End Function

Private Sub Z()
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
BrwDy DclLinCntzP(CPj)
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
DclLinCnt = IxoPrvCdLin(Src, Fm) + 1
End Function

Function IxoPrvCdLin&(Src$(), Fm)
Dim O&
For O = Fm - 1 To 0 Step -1
    If IsLinCd(Src(O)) Then IxoPrvCdLin = O: Exit Function
Next
IxoPrvCdLin = -1
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
        DclDiczP.Add MdDn(M), Dcl
    End If
Next
End Function

Function DclItr(M As CodeModule)
Asg Itr(DclLyzM(M)), DclItr
End Function

Function DclL$(Src$())
':DclL: :Lines ! comes fm a module
DclL = JnCrLf(DclLy(Src))
End Function

Function DclLy(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim N&, O$()
   N = DclLinCnt(Src)
If N <= 0 Then Exit Function
O = FstNEle(Src, N)
DclLy = O
'Brw LyzNNAp("N Src DclLy", N, AddIxPfx(Src), O): Stop
End Function

Function LineszMLC$(M As CodeModule, Lno&, Cnt&)
If Lno <= 0 Then Exit Function
If Cnt <= 0 Then Exit Function
If Lno > M.CountOfLines Then Exit Function
LineszMLC = M.Lines(Lno, Cnt)
End Function

Private Sub Z_DclzM()
Dim O$(), C As VBComponent
For Each C In CPj.VBComponents
    PushNB O, DclzM(C.CodeModule)
Next
VcLinesAy O
End Sub

Function DclzM$(M As CodeModule)
DclzM = LineszRTrim(LineszMLC(M, 1, DclLinCntzM(M)))
End Function

Function DclLyzM(M As CodeModule) As String()
DclLyzM = SplitCrLf(DclzM(M))
End Function

'
