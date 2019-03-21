Attribute VB_Name = "MVb_Itp"
Option Explicit
Function IntozItp(OInto, Itr, P) As String()
Dim O: O = OInto
Dim Obj
For Each Obj In Itr
    Push O, ObjPrp(Obj, P, eeNoThwNoInf)
Next
IntozItp = O
End Function
Function SyzItp(Itr, P) As String()
SyzItp = IntozItp(EmpSy, Itr, P)
End Function

