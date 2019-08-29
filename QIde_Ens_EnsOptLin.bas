Attribute VB_Name = "QIde_Ens_EnsOptLin"
Option Explicit
Option Compare Text
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_OptLin."
Const LExplicit$ = "Option Explicit"
Const LCmpBin$ = "Option Compare Binary"
Const LCmpDb$ = "Option Compare Database"
Const LCmpTxt$ = "Option Compare Text"

Sub EnsOptLinP()
EnsOptLinzP CPj
End Sub

Sub EnsOptLinM()
EnsOptLinzM CMd
End Sub

Private Sub EnsOptLinzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsOptLinzM C.CodeModule
Next
End Sub
Private Sub Z_EnsOptLinzM()
Dim Md As CodeModule
Const Mdn$ = "AA"
GoSub Setup
GoSub T0
GoSub Clean
Exit Sub
T0:
    Set Md = MdzDn(Mdn)
    GoTo Tst
Tst:
    EnsOptLinzM Md
    Return
Setup:
    AddCls Mdn
    Return
Clean:
    RmvMd Mdn
    Return
End Sub

Private Sub EnsOptLinzM(M As CodeModule)
If IsMdEmp(M) Then Exit Sub
DltOptLin M, LCmpDb
DltOptLin M, LCmpBin
EnsOptLin M, LCmpTxt
EnsOptLin M, LExplicit
End Sub

Private Sub Z_LnoAftOptqImpl()
Dim Md As CodeModule
GoSub T0
Exit Sub
T0:
    Set Md = MdzDn("ATaxExpCmp_OupTblGenr")
    Ept = 2&
    GoTo Tst
Tst:
    Act = LnoAftOptqImpl(Md)
    C
    Return
End Sub

Function IxoAftOptqImplzS&(Src$())
Dim Fnd As Boolean, J%, IsOpt As Boolean, L$
For J = 0 To UB(Src)
    L = Src(J)
    'IsOpt = IsLin_OfOpt_OrImpl_OrBlnk(L)
    Select Case True
    Case Fnd And IsOpt:
    Case Fnd: IxoAftOptqImplzS = J: Exit Function
    Case IsOpt: Fnd = True
    End Select
Next
IxoAftOptqImplzS = J
End Function

Function LnoAftOptqImpl&(M As CodeModule)
Dim Fnd As Boolean, J%, IsOpt As Boolean, L$
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    'IsOpt = IsLin_OfOpt_OrImpl_OrBlnk(L)
    Select Case True
    Case Fnd And IsOpt:
    Case Fnd: LnoAftOptqImpl = J: Exit Function
    Case IsOpt: Fnd = True
    End Select
Next
LnoAftOptqImpl = J
End Function

Private Function OptLno%(M As CodeModule, OptLin)
Dim J&
For J = 1 To M.CountOfDeclarationLines
   If M.Lines(J, 1) = OptLin Then OptLno = J: Exit Function
Next
End Function

Private Sub EnsOptLin(M As CodeModule, OptLin)
Const CSub$ = CMod & "EnsOptLin"
If M.CountOfLines = 0 Then Exit Sub
If OptLno(M, OptLin) > 0 Then Exit Sub
M.InsertLines 1, OptLin
InfLin CSub, "[" & OptLin & "] is Inserted", "Md", Mdn(M)
End Sub

Private Sub DltOptLin(M As CodeModule, OptLin)
Const CSub$ = CMod & "DltOptLin"
Dim I%: I = OptLno(M, OptLin)
If I = 0 Then Exit Sub
M.DeleteLines I
Inf CSub, "[" & OptLin & "] line is deleted", "Md Lno", Mdn(M), I
End Sub

Private Sub Z()
QIde_Ens_EnsOptLin:
End Sub

'
