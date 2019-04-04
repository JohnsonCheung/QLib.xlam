Attribute VB_Name = "MVb_Ay_Op_Rpl"
Option Explicit
Function AyRplFTIx(Ay, B As FTIx, ByAy)
Dim X As AyABC
Set X = AyabCzFTIx(Ay, B)
AyRplFTIx = AyAddAp(X.A, ByAy, X.C)
End Function
