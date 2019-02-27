Attribute VB_Name = "MVb_Ay_Op_ReOrder"
Option Explicit

Function AyReOrd(Ay, SubAy)
Dim HasSubAy: HasSubAy = AyIntersect(Ay, SubAy)
Dim Rest: Rest = AyMinus(Ay, SubAy)
AyReOrd = AyAdd(HasSubAy, Rest)
End Function
