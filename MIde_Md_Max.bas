Attribute VB_Name = "MIde_Md_Max"
Option Explicit


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


