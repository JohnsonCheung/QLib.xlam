Attribute VB_Name = "QIde_MthId"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_MthId."
Private Const Asm$ = "QIde"
':MthMLin = "M for Modified.  Fmt is [ShtMthMdy.ShtMthKd.MthnRst].  MthnRst is MthMLin with MthMdy and MthTy removed."
':MthSrtKey$ = "It is Str.  Fmt is [MthMdy:Mthn]"
Function MthSrtKey$(ShtMthMdy$, Mthn)
MthSrtKey = ShtMthMdy & ":" & Mthn
End Function

Function MthSrtKeyzLin(MthLin) ' MthKey is Mdy.Nm
With Mthn3zL(MthLin)
MthSrtKeyzLin = .ShtMdy & "." & .Nm
End With
End Function
