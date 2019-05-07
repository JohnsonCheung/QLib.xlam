Attribute VB_Name = "QIde_Md_Max"
Option Explicit
Private Const CMod$ = "MIde_Md_Max."
Private Const Asm$ = "QIde"


Function MaxLinCntMdNm$()
MaxLinCntMdNm = MaxLinCntMdNmzPj(CurPj)
End Function

Function MaxLinCntMdNmzPj$(A As VBProject)
MaxLinCntMdNmzPj = MdNm(MaxLinCntMd(A))
End Function
Function MaxLinCntMd(A As VBProject) As CodeModule
Dim C As VBComponent, M&, N&, I
For Each C In A.VBComponents
    N = C.CodeModule.CountOfLines
    If N > M Then
        M = N
        Set MaxLinCntMd = C.CodeModule
    End If
Next
End Function
Function CvMd(A) As CodeModule
Set CvMd = A
End Function


