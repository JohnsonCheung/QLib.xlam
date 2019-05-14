Attribute VB_Name = "QIde_Ens_OptLin"
Option Explicit
Option Compare Text
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_OptLin."
Const OptLinOfExplicit$ = "Option Explicit"
Const OptLinOfCmpBin$ = "Option Compare Binary"
Const OptLinOfCmpDb$ = "Option Compare Database"
Const OptLinOfCmpTxt$ = "Option Compare Text"

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
Private Sub Z_EnsOptLinOfMd()
Dim Md As CodeModule
Const Mdn = "AA"
GoSub Setup
GoSub T0
GoSub Clean
Exit Sub
T0:
    Set Md = MdzDNm(Mdn)
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
Private Sub EnsOptLinzM(A As CodeModule)
If IsMdOfEmp(A) Then Exit Sub
DltOptLin A, OptLinOfCmpDb
DltOptLin A, OptLinOfCmpBin
EnsOptLin A, OptLinOfCmpTxt
EnsOptLin A, OptLinOfExplicit
End Sub

Private Sub Z_LnoOf_AftOpt_AndImpl()
Dim Md As CodeModule
GoSub T0
Exit Sub
T0:
    Set Md = MdzDNm("ATaxExpCmp_OupTblGenr")
    Ept = 2&
    GoTo Tst
Tst:
    Act = LnoOf_AftOpt_AndImpl(Md)
    C
    Return
End Sub
Function IxOfAftOptAndImplzS&(Src$())
Dim Fnd As Boolean, J%, IsOpt As Boolean, L$
For J = 0 To UB(Src)
    L = Src(J)
    IsOpt = IsLin_OfOpt_OrImpl_OrBlank(L)
    Select Case True
    Case Fnd And IsOpt:
    Case Fnd: IxOfAftOptAndImplzS = J: Exit Function
    Case IsOpt: Fnd = True
    End Select
Next
IxOfAftOptAndImplzS = J
End Function

Function LnoOf_AftOpt_AndImpl&(A As CodeModule)
Dim Fnd As Boolean, J%, IsOpt As Boolean, L$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    IsOpt = IsLin_OfOpt_OrImpl_OrBlank(L)
    Select Case True
    Case Fnd And IsOpt:
    Case Fnd: LnoOf_AftOpt_AndImpl = J: Exit Function
    Case IsOpt: Fnd = True
    End Select
Next
LnoOf_AftOpt_AndImpl = J
End Function

Private Function OptLno%(A As CodeModule, OptLin)
Dim J&
For J = 1 To A.CountOfDeclarationLines
   If A.Lines(J, 1) = OptLin Then OptLno = J: Exit Function
Next
End Function

Private Sub EnsOptLin(A As CodeModule, OptLin)
Const CSub$ = CMod & "EnsOptLin"
If A.CountOfLines = 0 Then Exit Sub
If OptLno(A, OptLin) > 0 Then Exit Sub
A.InsertLines 1, OptLin
InfLin CSub, "[" & OptLin & "] is Inserted", "Md", Mdn(A)
End Sub

Private Sub DltOptLin(A As CodeModule, OptLin)
Const CSub$ = CMod & "DltOptLin"
Dim I%: I = OptLno(A, OptLin)
If I = 0 Then Exit Sub
A.DeleteLines I
Inf CSub, "[" & OptLin & "] line is deleted", "Md Lno", Mdn(A), I
End Sub
