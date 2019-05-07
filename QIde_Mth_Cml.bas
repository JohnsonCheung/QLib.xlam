Attribute VB_Name = "QIde_Mth_Cml"
Option Explicit
Private Const CMod$ = "MIde_Mth_Cml."
Private Const Asm$ = "QIde"
#Const Sav = True
Public Const DoczMthCml$ = "NewType:Sy."

Function MthCmlAsetInPj(Optional WhStr$) As Aset
Set MthCmlAsetInPj = CmlAset(MthNsetP(WhStr).Sy)
End Function

Function MthCmlFny(NDryCol%) As String()
MthCmlFny = AddAyAp(SyzSsLin("Mdy Kd Mth"), FnyzPfxN("Seg", NDryCol - 3))
End Function
Function MthCmlWs(Optional Vis As Boolean) As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = MthCmlLinWsBase
Set Lo = FstLo(Ws)
AddFml Lo, "Sel", "" ' "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
CrtLozAyH Seg1ErNy, WbzLo(Lo), "Seg1Er"
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

