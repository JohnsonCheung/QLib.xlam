Attribute VB_Name = "MIde_Exp_SrcPth_Pjf"
Option Explicit
Function Fxa$(FxaNm, SrcPth)
Fxa = DistPth(SrcPth) & FxaNm & ".xlam"
End Function

Function Fba$(FbaNm, SrcPth)
Fba = PthEns(SrcPth & "Dist") & FbaNm & ".accdb"
End Function

