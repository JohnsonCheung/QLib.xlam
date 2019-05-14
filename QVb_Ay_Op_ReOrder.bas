Attribute VB_Name = "QVb_Ay_Op_ReOrder"
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_ReOrder."
Private Const Asm$ = "QVb"

Function AyReOrd(Ay, SubAy)
Dim HasSubAy: HasSubAy = IntersectAy(Ay, SubAy)
Dim Rest: Rest = MinusAy(Ay, SubAy)
AyReOrd = AddAy(HasSubAy, Rest)
End Function
