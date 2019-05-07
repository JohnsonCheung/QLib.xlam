Attribute VB_Name = "QIde_Bld_BldFxa"
Option Explicit
Private Const CMod$ = "BCrtFxa."
Private Type A
    Pj As VBProject
End Type
Private Type R1
    Srcp As String
    DistFxa As String
    RfSrc() As String:      End Type
Private Type R2
    DistPj As VBProject:    End Type
Private A As A
Private R1 As R1
Private R2 As R2
Sub BldFxaP()
BldFxazPj CurPj
End Sub

Sub BldFxazPj(Pj As VBProject)
Set A.Pj = Pj
R1_Exp: Nop
    ExpPj A.Pj
    R1.Srcp = Srcp(A.Pj)
    R1.DistFxa = DistFxa(R1.Srcp)
    R1.RfSrc = RfSrczSrcp(R1.Srcp)
R2_CrtFxa: Nop
    CrtFxa R1.DistFxa
    Set R2.DistPj = PjzFxa(R1.DistFxa)
R3_AddRf: Nop
    AddRfzSrc R2.DistPj, R1.RfSrc
R4_LoadBas: Nop
    LoadBas R2.DistPj, R1.Srcp
End Sub
End Sub
Private Sub ZZ()
QIde_Bld_BldFxa:
End Sub
