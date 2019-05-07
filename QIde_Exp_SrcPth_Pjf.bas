Attribute VB_Name = "QIde_Exp_SrcPth_Pjf"
Option Explicit
Private Const CMod$ = "MIde_Exp_SrcPth_Pjf."
Private Const Asm$ = "QIde"
Function Fxa$(FxaNm, Srcp)
Fxa = Distp(Srcp) & FxaNm & ".xlam"
End Function

Function Fba$(FbaNm, Srcp)
Fba = EnsPth(Srcp & "Dist") & FbaNm & ".accdb"
End Function

