Attribute VB_Name = "QIde_Bld_BldFxa"
Option Explicit
Private Const CMod$ = "BCrtFxa."
Sub GenFxaP()
GenFxazP CPj
End Sub

Sub GenFxazP(Pj As VBProject)
Dim P$:                     P = Srcp(Pj)
Dim OFxa$:               OFxa = DistFxa(P)
                                ExpPj Pj
                                CrtFxa OFxa
Dim OPj As VBProject: Set OPj = PjzFxa(OFxa)
                                AddRfzSrc OPj, RfSrczSrcp(P)
                                LoadBas OPj, P
End Sub
Private Sub ZZ()
QIde_Bld_GenFxa:
End Sub
