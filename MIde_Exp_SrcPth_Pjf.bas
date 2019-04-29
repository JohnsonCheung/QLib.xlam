Attribute VB_Name = "MIde_Exp_SrcPth_Pjf"
Option Explicit
Function Fxa$(FxaNm, Srcp)
Fxa = DistPth(Srcp) & FxaNm & ".xlam"
End Function

Function Fba$(FbaNm, Srcp)
Fba = EnsPth(Srcp & "Dist") & FbaNm & ".accdb"
End Function

