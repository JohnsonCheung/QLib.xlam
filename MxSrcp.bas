Attribute VB_Name = "MxSrcp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxSrcp."
Function SrcpPj$()
SrcpPj = Srcp(CPj)
End Function

Function SrcpzCmp$(A As VBComponent)
SrcpzCmp = Srcp(PjzC(A))
End Function

Function SrcpzPjf$(Pjf)
SrcpzPjf = EnsPth(Pjf & ".src")
End Function

Sub EnsSrcp(P As VBProject)
EnsPthAll Srcp(P)
End Sub

Function SrcpzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
SrcpzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Function SrcpP$()
SrcpP = Srcp(CPj)
End Function

Function Srcp$(P As VBProject)
Srcp = SrcpzPjf(Pjf(P))
End Function

