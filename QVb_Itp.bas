Attribute VB_Name = "QVb_Itp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Itp."
Private Const Asm$ = "QVb"
Function IntozItrP(OInto, Itr, P As PrpPth) As String()
Dim O: O = OInto
Dim Obj As Object
For Each Obj In Itr
    Push O, Prp(Obj, P, EiNoThwQuiet)
Next
IntozItrP = O
End Function
Function SyzItrP(Itr, P As PrpPth) As String()
SyzItrP = IntozItrP(EmpSy, Itr, P)
End Function

