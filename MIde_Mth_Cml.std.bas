Attribute VB_Name = "MIde_Mth_Cml"
Option Explicit
Function MthCmlAsetPj(Optional WhStr$) As Aset
Set MthCmlAsetPj = CmlAset(JnSpc(MthAsetPj(WhStr).Itms))
End Function


Function MthCmlFny(NDryCol%) As String()
MthCmlFny = AyAddAp(SySsl("Mdy Kd Mth"), FnyzPfxN("Seg", NDryCol - 3))
End Function
Function MthCmlWs(Optional Vis As Boolean) As Worksheet
Dim Ws As Worksheet
Dim Lo As ListObject
Set Ws = MthCmlWsBase
Set Lo = FstLo(Ws)
AddFml Lo, "Sel", "=IF(ISNA(VLOOKUP([@Seg1],Seg1Er,1,True))),"""",""Err"")"
LozAyH Seg1ErNy, WbLo(Lo), "Seg1Er"
Lo.Application.Visible = Vis
Set MthCmlWs = Lo.Parent
End Function
Function MthCmlWsBase(Optional Vis As Boolean) As Worksheet
Dim Dry()
Dry = DryzSslAy(MthCmlVbe)
Set MthCmlWsBase = WszDrs(Drs(MthCmlFny(NColDry(Dry)), Dry), Vis:=Vis)
End Function

Sub MthCmlAyVbeBrw()
Brw AyAlign3T(MthCmlAyVbe)
End Sub
Function MthCmlVbe() As String()
MthCmlVbe = MthCmlAyzVbe(CurVbe)
End Function
Function MthCmlAyVbe() As String()
MthCmlAyVbe = MthCmlAyzVbe(CurVbe)
End Function
Function MthCmlAyzVbe(A As Vbe) As String()
Dim L
For Each L In Itr(MthDNyVbe(A))
    PushI MthCmlAyzVbe, MthCml(L)
Next
End Function

Function MthCml$(MthDNm)
Dim Ay$(): Ay = SplitDot(MthDNm)
Dim Pub$: If Ay(2) = "" Then Pub = ". " Else Pub = Ay(2) & " "
Dim Kd$: Kd = Ay(1) & " "
MthCml = Pub & Kd & Cmlxx(Ay(0))
End Function


