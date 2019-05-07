Attribute VB_Name = "QVb_Ay_FmTo_To_Dry"
Option Explicit
Private Const CMod$ = "MVb_Ay_FmTo_To_Dry."
Private Const Asm$ = "QVb"

Function DryzAddAyC(Ay, C) As Variant()
'XCDry is AyItmX-Const-Dry
Dim I
For Each I In Itr(Ay)
    PushI DryzAddAyC, Array(I, C)
Next
End Function

Function DryzCAddAy(A, C) As Variant()
'CXDry is Const-AyItmX-Dry
Dim I
For Each I In Itr(A)
    PushI DryzCAddAy, Array(C, I)
Next
End Function

Function DryzAyzTyNmVal(Ay) As Variant()
Dim I
For Each I In Itr(Ay)
    PushI DryzAyzTyNmVal, Array(TypeName(I), I)
Next
End Function

Sub DmpAyzTyNmVal(Ay)
DmpDry DryzAyzTyNmVal(Ay)
End Sub
