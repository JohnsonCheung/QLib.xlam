Attribute VB_Name = "QVb_Itp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Itp."
Private Const Asm$ = "QVb"
Function IntozItrP(OInto, Itr, PrpPth, Optional ThwEr As EmThw) As String()
Dim O: O = OInto
Dim Obj As Object
For Each Obj In Itr
    Push O, Prp(Obj, PrpPth, ThwEr)
Next
IntozItrP = O
End Function
Function SyzItrP(Itr, PrpPth) As String()
SyzItrP = IntozItrP(EmpSy, Itr, PrpPth)
End Function

