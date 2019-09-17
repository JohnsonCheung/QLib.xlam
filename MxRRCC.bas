Attribute VB_Name = "MxRRCC"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRRCC."
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
Type RC
    R As Long
    C As Long
End Type
Function HasRC(A As RRCC, B As RC) As Boolean
If NBet(B.R, A.R1, A.R2) Then Exit Function
If NBet(B.C, A.C1, A.C2) Then Exit Function
HasRC = True
End Function
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

Function IsEqRRCC(A As RRCC, B As RRCC) As Boolean
Dim A1 As RRCC: A1 = NormRRCC(A)
Dim B1 As RRCC: B1 = NormRRCC(B)
If A1.R1 <> B1.R1 Then Exit Function
If A1.R2 <> B1.R2 Then Exit Function
If A1.C1 <> B1.C1 Then Exit Function
If A1.C2 <> B1.C2 Then Exit Function
IsEqRRCC = True
End Function

Function NormRRCC(A As RRCC) As RRCC
Dim O As RRCC
With O
    If A.R1 < 0 Then .R1 = 0
    If A.R2 < 0 Then .R2 = 0
    If A.C1 < 0 Then .C1 = 0
    If A.C2 < 0 Then .C2 = 0
    If .R1 > .R2 Then .R1 = 0: .R2 = 0
End With
End Function

Function RRCCIsEmp(A As RRCC) As Boolean
RRCCIsEmp = IsEqRRCC(A, EmpRRCC)
End Function
Function EmpRRCC() As RRCC
End Function
Function RRCCTy(A As RRCC) As EmRRCCTy

End Function
Function RRCCLin$(A As RRCC)
Dim O$
'Select Case RRCCTy(A)
'Case EiRCC
'    O = FmtQQ("RCC(? ? ?) ", R1, C1, C2)
'Case EiRR
'    O = FmtQQ("RR(? ?) ", R1, R2)
'Case EiRow
'    O = FmtQQ("R(?)", R1)
'Case Else
'    'Thw CSub TpPos_FmtStr", "Invalid {TpPos}", A.Ty
'End Select
End Function
