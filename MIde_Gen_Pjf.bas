Attribute VB_Name = "MIde_Gen_Pjf"
Option Explicit
Const TermzSrcRoot$ = "It a Pth with Fdr eq '.src'"
Public Const DocOfSrcp$ = "Src-Pth:Src.p. It a Pth with Fdr is PjFn and ParFdr is SrcRoot"
Public Const DocOfDistRoot$ = "It a Pth with Fdr eq '.dist'"
Public Const DocOfDistPth$ = "It an InstPth with Fdr is InstPth and ParFdr is DistRoot"
Public Const DocOfPthInst$ = "It a InstFdr of under a given Pth"
Public Const DocOfInstPth$ = "It a Pth with Fdr is InstNm"

Function SrcRoot$(Srcp$)
SrcRoot = ParPth(Srcp)
End Function

Function DistPth$(Srcp$)
DistPth = AddFdrEns(ParPth(Srcp), ".Dist")
End Function

Function DistFba$(Srcp$)
DistFba = DistPjf(Srcp, ".accdb")
End Function

Private Function DistPjf(Srcp$, Ext$)
Dim P$: P = DistPth(Srcp)
DistPjf = DistPth(Srcp) & RplExt(Fdr(ParPth(P)), Ext)
End Function

Function DistFxa$(Srcp$)
DistFxa = DistPjf(Srcp, ".xlam")
End Function

Function DistFxazNxt$(Srcp$)
DistFxazNxt = NxtFfn(DistFxa(Srcp))
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

