Attribute VB_Name = "QTp_Tp_RRCC"
Option Explicit
Private Const CMod$ = "MTp_Tp_RRCC."
Private Const Asm$ = "QTp"

Function IsEmpRRCC(A As RRCC) As Boolean
End Function

Function CvRRCC(A) As RRCC
Set CvRRCC = A
End Function
Function NewRRCC(R1, R2, C1, C2) As RRCC
Set NewRRCC = New RRCC
With NewRRCC
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function
