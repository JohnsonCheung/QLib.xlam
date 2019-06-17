Attribute VB_Name = "QVb_Ay_FmTo_To_Dry"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_FmTo_To_Dry."
Private Const Asm$ = "QVb"

Function DryzAyzAddC(Ay, C) As Variant()
'XCDry is AyItmX-Const-Dry
Dim I
For Each I In Itr(Ay)
    PushI DryzAyzAddC, Array(I, C)
Next
End Function

Function DryzCAyzAdd(A, C) As Variant()
'CXDry is Const-AyItmX-Dry
Dim I
For Each I In Itr(A)
    PushI DryzCAyzAdd, Array(C, I)
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
