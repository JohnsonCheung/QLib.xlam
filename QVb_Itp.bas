Attribute VB_Name = "QVb_Itp"
Option Explicit
Private Const CMod$ = "MVb_Itp."
Private Const Asm$ = "QVb"
Function IntozItrP(OInto, Itr, P$) As String()
Dim O: O = OInto
Dim Obj As Object
For Each Obj In Itr
    Push O, Prp(Obj, P, EiNoThwQuiet)
Next
IntozItrP = O
End Function
Function SyzItrP(Itr, P$) As String()
SyzItrP = IntozItrP(EmpSy, Itr, P)
End Function

