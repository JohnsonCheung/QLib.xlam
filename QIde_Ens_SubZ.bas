Attribute VB_Name = "QIde_Ens_SubZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZ."
Private Const Asm$ = "QIde"

Function CdSubZzM$(M As CodeModule)
'SubZ is [Mth-`Sub Z()`-Lines], each line is calling a Z_XX, where Z_XX is a testing function
Dim A As Drs: A = ColEq(DMthP, "Mdn", Mdn(M))
Dim B As Drs: B = ColPfx(A, "Mthn", "Z_")
Dim Mthny$(): Mthny = StrCol(A, "Mthn")
Dim S$(): S = SrtAy(Mthny)
Erase XX
X "Private Sub ZZ()"
X S
X "End Sub"
CdSubZzM = JnCrLf(XX)
End Function

