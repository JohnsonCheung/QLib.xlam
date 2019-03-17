Attribute VB_Name = "MVb_Ay_Op_Rpl"
Function AyRplFTIx(Ay, B As FTIx, ByAy)
Dim X As AyABC
Set X = AyABCzFTIx(Ay, B)
AyRplFTIx = AyAddAp(X.A, ByAy, X.C)
End Function
