Attribute VB_Name = "MIde_Mth_Cnt"
Option Explicit

Function NMthzMd%(A As CodeModule, Optional WhStr$)
NMthzMd = NMthzSrc(Src(A), WhStr)
End Function

Function NSrcLinPj&(A As VBProject)
Dim O&, C As VBComponent
For Each C In A.VBComponents
    O = O + C.CodeModule.CountOfLines
Next
NSrcLinPj = O
End Function

Function NPubMthMd%(A As CodeModule)
NPubMthMd = NMthzSrc(Src(A), "-Pub")
End Function
Function NPubMthVbe%(A As Vbe)
Dim O%, P As VBProject
For Each P In A.VBProjects
    O = O + NPubMthPj(P)
Next
NPubMthVbe = O
End Function
Property Get NPubMth%()
NPubMth = NPubMthVbe(CurVbe)
End Property

Function NPubMthPj%(A As VBProject)
Dim O%, C As VBComponent
For Each C In A.VBComponents
    O = O + NPubMthMd(C.CodeModule)
Next
NPubMthPj = O
End Function

Function NMthzSrc%(A$(), Optional WhStr$)
NMthzSrc = Sz(MthIxAy(A, WhStr$))
End Function
