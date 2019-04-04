Attribute VB_Name = "MIde_Gen_Pjf"
Option Explicit
Const TermzSrcRoot$ = "It a Pth with Fdr eq '.src'"
Public Const DocOfSrcp$ = "Src-Pth:Src.p. It a Pth with Fdr is PjFn and ParFdr is SrcRoot"
Public Const DocOfDistRoot$ = "It a Pth with Fdr eq '.dist'"
Public Const DocOfDistPth$ = "It an InstPth with Fdr is InstPth and ParFdr is DistRoot"
Public Const DocOfPthInst$ = "It a InstFdr of under a given Pth"
Public Const DocOfInstPth$ = "It a Pth with Fdr is InstNm"
Function SrcRoot$(Srcp)
SrcRoot = ParPth(Srcp)
End Function
Function DistPth$(Srcp)
DistPth = AddFdrEns(PthUp(Srcp, 2), ".dist", Fdr(Srcp))
End Function

Function DistFba$(Srcp)
DistFba = DistPth(Srcp) & RplExt(Fdr(Srcp), ".accdb")
End Function

Function DistFxa$(Srcp)
DistFxa = DistPth(Srcp) & RplExt(Fdr(Srcp), ".xlam")
End Function

Function DistFxazNxt$(Srcp)
DistFxazNxt = NxtFfn(DistFxa(Srcp))
End Function

Private Sub Z_DistPjf()
Dim Pth
For Each Pth In SrcpAyzExpgzInst
    Debug.Print Pth
    Debug.Print DistFba(Pth)
    Debug.Print DistFxa(Pth)
    Debug.Print
Next
End Sub

