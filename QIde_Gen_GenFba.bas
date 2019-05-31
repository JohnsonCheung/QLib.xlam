Attribute VB_Name = "QIde_Gen_GenFba"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Fba."
Private Const Asm$ = "QIde"
Sub GenFbaP()
GenFbazP CPj
End Sub
Sub GenFbazP(Pj As VBProject)
'Exp
Dim P$
Dim OFba$
       ExpPj Pj
   P = Srcp(Pj)
OFba = DistFba(P)
'CrtFba
Dim Acs As New Access.Application:
Dim DistPj As VBProject:
    'DltFfnIf Fba
    CrtFb OFba
    OpnFb Acs, OFba
Set DistPj = _
    PjzAcs(Acs)
    'AddRfzDistPj DistPj
'LoadBas
    'LoadBas DistPj
'LoadFrm
Dim FFfny$()
    Dim FrmFfn, I, N$
    For Each FrmFfn In Itr(FrmFfny(P))
        N = RmvExt(RmvExt(FrmFfn))
        Acs.LoadFromText acForm, RmvExt(RmvExt(FrmFfn)), FrmFfn
    Next
'QUit
    QuitAcs Acs
End Sub

