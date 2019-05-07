Attribute VB_Name = "QIde_Ens_OptLin"
Option Explicit
Option Compare Text
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_OptLin."
Const OptLinzExplicit$ = "Option Explicit"
Const OptLinzCmpBin$ = "Option Compare Binary"
Const OptLinzCmpDb$ = "Option Compare Database"
Const OptLinzCmpTxt$ = "Option Compare Text"

Sub EnsOptLinPj()
EnsOptLinzPj CurPj
End Sub

Sub EnsMdOptLin()
EnsOptLinzMd CurMd
End Sub

Private Sub EnsOptLinzPj(Pj As VBProject)
Dim C As VBComponent
For Each C In Pj.VBComponents
    EnsOptLinzMd C.CodeModule
Next
End Sub
Private Sub Z_EnsOptLinzMd()
Dim Md As CodeModule
Const MdNm$ = "AA"
GoSub Setup
GoSub T0
GoSub Clean
Exit Sub
T0:
    Set Md = MdzDNm(MdNm)
    GoTo Tst
Tst:
    EnsOptLinzMd Md
    Return
Setup:
    CrtCls MdNm
    Return
Clean:
    RmvMd MdNm
    Return
End Sub
Private Sub EnsOptLinzMd(A As CodeModule)
RmvOptLin A, OptLinzCmpDb
RmvOptLin A, OptLinzCmpBin
EnsOptLin A, OptLinzCmpTxt
EnsOptLin A, OptLinzExplicit
End Sub
Private Sub EnsCLibzPj(A As VBProject, Optional B As EmLibNmTy)
If A.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In A.VBComponents
    EnsCLib C.CodeModule, B
Next
End Sub
Private Sub Z_LnozAftOptzAndImp()
Dim Md As CodeModule
GoSub T0
Exit Sub
T0:
    Set Md = MdzDNm("ATaxExpCmp_OupTblGenr")
    Ept = 2&
    GoTo Tst
Tst:
    Act = LnozAftOptzAndImpl(Md)
    C
    Return
End Sub
Function LnozAftOptzAndImpl&(A As CodeModule)
Dim Fnd As Boolean, J%, IsOpt As Boolean, L$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    IsOpt = IsOptLinzOrImplzOrBlank(L)
    Select Case True
    Case Fnd And IsOpt:
    Case Fnd: LnozAftOptzAndImpl = J: Exit Function
    Case IsOpt: Fnd = True
    End Select
Next
LnozAftOptzAndImpl = J
End Function

Private Function OptLno%(A As CodeModule, OptLin$)
Dim J&
For J = 1 To A.CountOfDeclarationLines
   If A.Lines(J, 1) = OptLin Then OptLno = J: Exit Function
Next
End Function

Private Sub EnsOptLin(A As CodeModule, OptLin$)
Const CSub$ = CMod & "EnsOptLin"
If A.CountOfLines = 0 Then Exit Sub
If OptLno(A, OptLin) > 0 Then Exit Sub
A.InsertLines 1, OptLin
InfLin CSub, "[" & OptLin & "] is Inserted", "Md", MdNm(A)
End Sub

Private Sub RmvOptLin(A As CodeModule, OptLin$)
Const CSub$ = CMod & "RmvOptLin"
Dim I%: I = OptLno(A, OptLin)
If I = 0 Then Exit Sub
A.DeleteLines I
Inf CSub, "[" & OptLin & "] line is deleted", "Md Lno", MdNm(A), I
End Sub
