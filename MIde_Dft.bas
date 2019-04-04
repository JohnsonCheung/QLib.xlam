Attribute VB_Name = "MIde_Dft"
Option Explicit
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
Function SiOfPj&()
SiOfPj = SizPj(CurPj)
End Function

Private Sub Z()
MIde__Dft:
End Sub
