Attribute VB_Name = "QVb_Cml_Rel"
Option Explicit
Private Const CMod$ = "MVb_Cml_Rel."
Private Const Asm$ = "QVb"
Function CmlRel(Ny$()) As Rel
Set CmlRel = Rel(CmlLy(Ny))
End Function
