Attribute VB_Name = "MVb_Itp"
Option Explicit
Function IntozItrP(OInto, Itr, P$) As String()
Dim O: O = OInto
Dim Obj As Object
For Each Obj In Itr
    Push O, Prp(Obj, P, jNoThwNoInf)
Next
IntozItrP = O
End Function
Function SyzItrP(Itr, P$) As String()
SyzItrP = IntozItrP(EmpSy, Itr, P)
End Function

