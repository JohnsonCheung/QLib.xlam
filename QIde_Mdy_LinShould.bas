Attribute VB_Name = "QIde_Mdy_LinShould"
Option Explicit
Private Const CMod$ = "MIde_Mdy_LinShould."
Private Const Asm$ = "QIde"
Function ShouldIns(IsUsing As Boolean, OldLin$, NewLin$) As Boolean
Select Case True
Case IsUsing And OldLin = "": ShouldIns = True
Case IsUsing And OldLin <> NewLin: ShouldIns = True
End Select
End Function

Function ShouldDlt(IsUsing As Boolean, OldLin$, NewLin$) As Boolean
Select Case True
Case IsUsing And OldLin <> "" And OldLin <> NewLin: ShouldDlt = True
Case Not IsUsing And OldLin <> "": ShouldDlt = True
End Select
End Function

