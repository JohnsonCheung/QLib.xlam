Attribute VB_Name = "MIde_Mdy_LinShould"
Option Explicit
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

