Attribute VB_Name = "MIde_Mth_Cml"
Option Explicit
#Const Sav = True
Public Const DocOfMthCml$ = "NewType:Sy."

Function MthCmlAsetInPj(Optional WhStr$) As Aset
Set MthCmlAsetInPj = CmlAset(MthNsetInPj(WhStr).Sy)
End Function

Function MthCmlFny(NDryCol%) As String()
MthCmlFny = AyAddAp(SySsl("Mdy Kd Mth"), FnyzPfxN("Seg", NDryCol - 3))
End Function
Function MthCmlWs(Optional Vis As Boolean) As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = MthCmlLinWsBase
Set Lo = FstLo(Ws)
AddFml Lo, "Sel", "" ' "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
LozAyH Seg1ErNy, WbzLo(Lo), "Seg1Er"
Lo.Application.Visible = Vis
Set MthCmlWs = Lo.Parent
End Function
Function MthCmlLinWsBase() As Worksheet
Dim Dry()
Dry = DryzSslAy(MthCmlLyInVbe)
Set MthCmlLinWsBase = WszDrs(Drs(MthCmlFny(NColzDry(Dry)), Dry))
End Function

Sub BrwMthCmlLyInVbe()
Brw FmtSyT3(MthCmlLyInVbe)
End Sub

Function MthCmlLyInVbe() As String()
MthCmlLyInVbe = MthCmlLyzVbe(CurVbe)
End Function

Function MthCmlLyzVbe(A As Vbe) As String()
MthCmlLyzVbe = CmlLy(MthNyzVbe(A))
End Function

