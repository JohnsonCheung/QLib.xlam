Attribute VB_Name = "MDta_S1S2"
Option Explicit
Function S1S2DrszAy(A() As S1S2) As Drs
Set S1S2DrszAy = Drs("S1 S2", S1S2AyDry(A))
End Function

Function S1S2AyDry(A() As S1S2) As Variant()
Dim J%
For J = 0 To UB(A)
   With A(J)
       PushI S1S2AyDry, Array(.S1, .S2)
   End With
Next
End Function
