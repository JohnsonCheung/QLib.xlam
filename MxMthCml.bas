Attribute VB_Name = "MxMthCml"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthCml."
#Const Sav = True
':MthCml$ = "NewType:Sy."

Function AsetzMthCmlP() As Aset
Set AsetzMthCmlP = CmlAset(MthnsetP.Sy)
End Function

Function FnyzMthCml(NDyCol%) As String()
FnyzMthCml = AddAyAp(SyzSS("Mdy Kd Mth"), FnyzPfxN("Seg", NDyCol - 3))
End Function

Function WsoMthCm() As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = WsoMthCmlBase
Set Lo = FstLo(Ws)
'AddFml Lo, "Sel", "" ' "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
LozAyH Seg1ErNy, WbzLo(Lo), "Seg1Er"
Set WsoMthCm = ShwWs(Lo.Parent)
End Function

Function WsoMthCmlBase() As Worksheet
Dim Dy()
'Dy = DyoSslAy(MthCmlssAyInVbe)
Set WsoMthCmlBase = WszDrs(Drs(FnyzMthCml(NColzDy(Dy)), Dy))
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
