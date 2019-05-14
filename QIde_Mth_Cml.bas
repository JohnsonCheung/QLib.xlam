Attribute VB_Name = "QIde_Mth_Cml"
Option Explicit
Private Const CMod$ = "MIde_Mth_Cml."
Private Const Asm$ = "QIde"
#Const Sav = True
Public Const DoczMthCml$ = "NewType:Sy."

Function AsetzMthCmlP(Optional WhStr$) As Aset
Set AsetzMthCmlP = CmlAset(MthnsetP(WhStr).Sy)
End Function

Function FnyzMthCml(NDryCol%) As String()
FnyzMthCml = AddAyAp(SyzSS("Mdy Kd Mth"), FnyzPfxN("Seg", NDryCol - 3))
End Function
Function WszMthCm(Optional Vis As Boolean) As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = MthCmlLinWsBase
Set Lo = FstLo(Ws)
'AddFml Lo, "Sel", "" ' "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
CrtLozAyH Seg1ErNy, WbzLo(Lo), "Seg1Er"
Lo.Application.Visible = Vis
Set WszMthCm = Lo.Parent
End Function
Function MthCmlLinWsBase() As Worksheet
Dim Dry()
'Dry = DryzSslAy(MthCmlLyInVbe)
Set MthCmlLinWsBase = WszDrs(Drs(FnyzMthCml(NColzDry(Dry)), Dry))
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

