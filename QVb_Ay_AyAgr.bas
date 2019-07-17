Attribute VB_Name = "QVb_Ay_AyAgr"
Option Explicit
Option Compare Text
Function SumAy#(NumAy)
Dim O#, V: For Each V In Itr(NumAy)
    O = O + V
Next
SumAy = O
End Function
Function MaxAy(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay(0)
Dim V: For Each V In Ay
    If V > O Then O = V
Next
MaxAy = O
End Function

Function MinAy(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay(0)
Dim V: For Each V In Ay
    If V < 0 Then O = V
Next
MinAy = O
End Function

Function MinAyzGT0(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay(0)
Dim V: For Each V In Ay
    If V > 0 Then
        If O = 0 Then
            O = V
        Else
            If V < O Then O = V
        End If
    End If
Next
MinAyzGT0 = O
End Function

Function LcAgrP() As Drs
LcAgrP = AgrzNum(LinCntP)
End Function

Function LinCntP() As Long()
LinCntP = LinCntzP(CPj)
End Function

Function LinCntzP(P As VBProject) As Long()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI LinCntzP, C.CodeModule.CountOfLines
Next
End Function

Function CntNo0&(NumAy)
Dim O&
Dim V: For Each V In Itr(NumAy)
    If V <> 0 Then O = O + 1
Next
CntNo0 = O
End Function

Function AgrzNum(NumAy) As Drs
'Ret : Agr Val ! where *Arg has Cnt Avg Max Min Sum
Dim ODy()
Dim Sum#: Sum = SumAy(NumAy)
Dim NNo0&: NNo0 = CntNo0(NumAy)
Dim N&: N = Si(NumAy)
Dim AvgAll#, AvgNo0#
If N <> 0 Then AvgAll = Sum / N
If NNo0 <> 0 Then AvgNo0 = Sum / NNo0

Push ODy, Array("CntNo0", NNo0)
Push ODy, Array("CntAll", N)
Push ODy, Array("AvgNo0", AvgNo0)
Push ODy, Array("AvgAll", AvgAll)
Push ODy, Array("Sum", Sum)
Push ODy, Array("Max", MaxAy(NumAy))
Push ODy, Array("Min", MinAy(NumAy))
Push ODy, Array("MinGT0", MinAyzGT0(NumAy))
AgrzNum = DrszFF("Agr Val", ODy)
End Function

