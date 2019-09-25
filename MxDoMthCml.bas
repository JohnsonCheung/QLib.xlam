Attribute VB_Name = "MxDoMthCml"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxMthCml."
#Const Sav = True
':MthCml$ = "NewType:Sy."

Function FnyzMthCml(NDyCol%) As String()
FnyzMthCml = AddAyAp(SyzSS("Mdy Kd Mth"), FnyzPfxN("Seg", NDyCol - 3))
End Function

Function DoMthCmlP() As Drs

End Function

Function WsoMthCmlP() As Worksheet
Set WsoMthCmlP = WszDrs(DoMthCmlP)
End Function

