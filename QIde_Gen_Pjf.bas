Attribute VB_Name = "QIde_Gen_Pjf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf."
Private Const Asm$ = "QIde"
Public Const DoctzSrcRoot$ = "It a Pth with Fdr eq '.src'"
Public Const DoctzSrcp$ = "Src-Pth:Src.p. It a Pth with Fdr is PjFn and ParFdr is SrcRoot"
Public Const DoctzDistp$ = "It an InstPth with Fdr is InstPth and ParFdr is DistRoot"
Public Const DoctzInstPth$ = "It a Pth with Fdr is InstNm"
Function SrcRoot$(Srcp$)
SrcRoot = ParPth(Srcp)
End Function
Function DistpP$() 'Distribution Path
DistpP = Distp(SrcpP)
End Function

Function Distp$(Srcp) 'Distribution Path
Distp = AddFdrEns(UpPth(Srcp, 2), ".Dist")
End Function

Function DistFba$(Srcp)
DistFba = DistPjf(Srcp, ".accdb")
End Function

Private Function DistPjf(Srcp, Ext) '
Dim P$:  P = Distp(Srcp)
Dim F1$: F1 = RplExt(Fdr(ParPth(P)), Ext)
Dim F2$: F2 = NxtFfnzNotIn(F1, PjfnAyV)
Dim F$:  F = NxtFfnzAva(P & F2)
DistPjf = F
End Function
Private Sub Z_DistFxa()
Dim Srcp$
GoSub T0
Exit Sub
T0:
    Srcp = SrcpP
    Ept = "C:\Users\user\Documents\Projects\Vba\QLib\.Dist\QLib(002).xlam"
    GoTo Tst
Tst:
    Act = DistFxa(Srcp)
    C
    Return
End Sub
Function DistFxa$(Srcp)
DistFxa = DistPjf(Srcp, ".xlam")
End Function

Private Sub Z_DistPjf()
Dim Pth, I
For Each I In SrcpSyOfExpgInst
    Pth = I
    Debug.Print Pth
    Debug.Print DistFba(Pth)
    Debug.Print DistFxa(Pth)
    Debug.Print
Next
End Sub

