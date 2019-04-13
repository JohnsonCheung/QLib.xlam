Attribute VB_Name = "MDta_S1S2"
Option Explicit
Function S1S2DrSumSi(A() As S1S2) As Drs
Set S1S2DrSumSi = Drs("S1 S2", S1S2AyDry(A))
End Function

Function S1S2AyDry(A() As S1S2) As Variant()
Dim J%
For J = 0 To UB(A)
   With A(J)
       PushI S1S2AyDry, Array(.s1, .s2)
   End With
Next
End Function
