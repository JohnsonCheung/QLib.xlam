Attribute VB_Name = "QIde_Gen_Pjf"
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf."
Private Const Asm$ = "QIde"
Const TermzSrcRoot$ = "It a Pth with Fdr eq '.src'"
Public Const DoczSrcp$ = "Src-Pth:Src.p. It a Pth with Fdr is PjFn and ParFdr is SrcRoot"
Public Const DoczDistRoot$ = "It a Pth with Fdr eq '.dist'"
Public Const DoczDistp$ = "It an InstPth with Fdr is InstPth and ParFdr is DistRoot"
Public Const DoczPthInst$ = "It a InstFdr of under a given Pth"
Public Const DoczInstPth$ = "It a Pth with Fdr is InstNm"
Function SrcRoot$(Srcp$)
SrcRoot = ParPth(Srcp)
End Function
Function DistpC$() 'Distribution Path
DistpC = Distp(SrcpC)
End Function

Function Distp$(Srcp$) 'Distribution Path
Distp = AddFdrEns(UpPth(Srcp, 2), ".Dist")
End Function

Function DistFba$(Srcp$)
DistFba = DistPjf(Srcp, ".accdb")
End Function

Private Function DistPjf(Srcp$, Ext$) '
Dim P$:  P = Distp(Srcp)
Dim F1$: F1 = RplExt(Fdr(ParPth(P)), Ext)
Dim F2$: F2 = NxtFfnzNotIn(F1, PjFnSyC)
Dim F$:  F = NxtFfnzAva(P & F2)
DistPjf = F
End Function
Private Sub Z_DistFxa()
Dim Srcp$
GoSub T0
Exit Sub
T0:
    Srcp = SrcpC
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = DistFxa(Srcp)
    C
    Return
End Sub
Function DistFxa$(Srcp$)
DistFxa = DistPjf(Srcp, ".xlam")
End Function

Private Sub Z_DistPjf()
Dim Pth$, I
For Each I In SrcpSyOfExpgInst
    Pth = I
    Debug.Print Pth
    Debug.Print DistFba(Pth)
    Debug.Print DistFxa(Pth)
    Debug.Print
Next
End Sub

