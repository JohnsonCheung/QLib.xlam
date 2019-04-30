Attribute VB_Name = "MDta_S1S2"
Option Explicit
Function S1S2DrSumSi(A As S1S2s) As Drs
Set S1S2DrSumSi = Drs("S1 S2", S1S2sDry(A))
End Function

Function S1S2sDry(A As S1S2s) As Variant()
Dim J%
For J = 0 To UB(A)
   With A(J)
       PushI S1S2sDry, Array(.S1, .S2)
   End With
Next
End Function
