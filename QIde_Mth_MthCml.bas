Attribute VB_Name = "QIde_Mth_MthCml"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Cml."
Private Const Asm$ = "QIde"
#Const Sav = True
':MthCml$ = "NewType:Sy."

Function AsetzMthCmlP() As Aset
Set AsetzMthCmlP = CmlAset(MthnsetP.Sy)
End Function

Function FnyzMthCml(NDyCol%) As String()
FnyzMthCml = AddAyAp(SyzSS("Mdy Kd Mth"), FnyzPfxN("Seg", NDyCol - 3))
End Function
Function WszMthCm() As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = MthCmlssWsBase
Set Lo = FstLo(Ws)
'AddFml Lo, "Sel", "" ' "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
LozAyH Seg1ErNy, WbzLo(Lo), "Seg1Er"
Set WszMthCm = ShwWs(Lo.Parent)
End Function
Function MthCmlssWsBase() As Worksheet
Dim Dy()
'Dy = DyoSslAy(MthCmlssAyInVbe)
Set MthCmlssWsBase = WszDrs(Drs(FnyzMthCml(NColzDy(Dy)), Dy))
End Function

Sub BrwMthCmlssAyV()
Brw FmtSy3Term(MthCmlssAyV)
End Sub

Function MthCmlssAyV() As String()
MthCmlssAyV = MthCmlssAyzV(CVbe)
End Function

Function MthCmlssAyzV(A As Vbe) As String()
MthCmlssAyzV = CmlssAy(MthNyzV(A))
End Function


'
