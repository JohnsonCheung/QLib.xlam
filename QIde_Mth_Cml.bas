Attribute VB_Name = "QIde_Mth_Cml"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Cml."
Private Const Asm$ = "QIde"
#Const Sav = True
Public Const DoczMthCml$ = "NewType:Sy."

Function AsetzMthCmlP() As Aset
Set AsetzMthCmlP = CmlAset(MthnsetP.Sy)
End Function

Function FnyzMthCml(NDyCol%) As String()
FnyzMthCml = AyzAddAp(SyzSS("Mdy Kd Mth"), FnyzPfxN("Seg", NDyCol - 3))
End Function
Function WszMthCm() As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = MthCmlLinWsBase
Set Lo = FstLo(Ws)
'AddFml Lo, "Sel", "" ' "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
LozAyH Seg1ErNy, WbzLo(Lo), "Seg1Er"
Set WszMthCm = ShwWs(Lo.Parent)
End Function
Function MthCmlLinWsBase() As Worksheet
Dim Dy()
'Dy = DyoSslAy(MthCmlLyInVbe)
Set MthCmlLinWsBase = WszDrs(Drs(FnyzMthCml(NColzDy(Dy)), Dy))
End Function

Sub BrwMthCmlLyV()
Brw FmtSy3Term(MthCmlLyV)
End Sub

Function MthCmlLyV() As String()
MthCmlLyV = MthCmlLyzV(CVbe)
End Function

Function MthCmlLyzV(A As Vbe) As String()
MthCmlLyzV = CmlLy(MthNyzV(A))
End Function

