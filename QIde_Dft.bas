Attribute VB_Name = "QIde_Dft"
Option Explicit
Private Const CMod$ = "MIde_Dft."
Private Const Asm$ = "QIde"
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CurMd
Else
   Set DftMd = A
End If
End Function

Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
   Set DftPj = CurPj
Else
   Set DftPj = A
End If
End Function


Function SizPj&(A As VBProject)
Dim O&, C As VBComponent
For Each C In A.VBComponents
    O = O + SizMd(C.CodeModule)
Next
SizPj = O
End Function
Function SiInPj&()
SiInPj = SizPj(CurPj)
End Function

Private Sub Z()
MIde__Dft:
End Sub
