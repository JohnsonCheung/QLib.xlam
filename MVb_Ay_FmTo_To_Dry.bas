Attribute VB_Name = "MVb_Ay_FmTo_To_Dry"
Option Explicit

Function DryzAyAddC(Ay, C) As Variant()
'XCDry is AyItmX-Const-Dry
Dim I
For Each I In Itr(Ay)
    PushI DryzAyAddC, Array(I, C)
Next
End Function

Function DryzCAddAy(A, C) As Variant()
'CXDry is Const-AyItmX-Dry
Dim I
For Each I In Itr(A)
    PushI DryzCAddAy, Array(C, I)
Next
End Function
