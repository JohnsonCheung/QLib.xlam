Attribute VB_Name = "QVb_Dta_RRCC"
Option Explicit
Enum EmRRCCTy
    EiRCC = 1
    EiRR = 2
    EiRow = 3
End Enum
Type RRCC
    R1 As Long 'all started from 1
    R2 As Long
    C1 As Long
    C2 As Long
End Type
Function RRCC(R1, R2, C1, C2) As RRCC
If R1 < 0 Then Exit Function
If R2 < 0 Then Exit Function
If C1 < 0 Then Exit Function
If C2 < 0 Then Exit Function
With RRCC
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function

Function IsRRCCzEmp(A As RRCC) As Boolean
IsRRCCEmp = True
With A
Select Case True
Case .R1 <= 0, .R2 <= 0, .C1 <= 0, .C2 <= 0
Case Else: IsRRCCzEmp = True
End Function
Function EmpRRCC() As RRCC
End Function
Function RRCCTy(A As RRCC) As EmRRCCTy

End Function
Property Get RRCCLin$(A As RRCC)
Dim O$
Select Case RRCCTy(A)
Case EiRCC
    O = FmtQQ("RCC(? ? ?) ", R1, C1, C2)
Case EiRR
    O = FmtQQ("RR(? ?) ", R1, R2)
Case EiRow
    O = FmtQQ("R(?)", R1)
Case Else
    'Thw CSub TpPos_FmtStr", "Invalid {TpPos}", A.Ty
End Select
End Property

