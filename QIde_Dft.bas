Attribute VB_Name = "QIde_Dft"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Dft."
Private Const Asm$ = "QIde"
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CMd
Else
   Set DftMd = A
End If
End Function

Function DftPj(P As VBProject) As VBProject
If IsNothing(P) Then
   Set DftPj = CPj
Else
   Set DftPj = P
End If
End Function


Function SizP&(P As VBProject)
Dim O&, C As VBComponent
For Each C In P.VBComponents
    O = O + SizMd(C.CodeModule)
Next
SizP = O
End Function
Function SiP&()
SiP = SizP(CPj)
End Function

Private Sub ZZ()
MIde__Dft:
End Sub
