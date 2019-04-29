Attribute VB_Name = "MVb_Ay_Op_Rpl"
Option Explicit
Function SyRplFTIx(Ay, B As FTIx, ByAy)
Dim X As AyABC
Set X = AyabcByFTIx(Ay, B)
SyRplFTIx = AyAddAp(X.A, ByAy, X.C)
End Function
