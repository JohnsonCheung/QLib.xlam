Attribute VB_Name = "QIde_Gen_GenFba"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Fba."
Private Const Asm$ = "QIde"
Sub GenFbaP()
GenFbazP CPj
End Sub
Sub GenFbazP(Pj As VBProject)
Dim Acs As New Access.Application, DistPj As VBProject
Dim Fba$:   Fba = DistFba(P)
                  DltFfnIf Fba
                  CrtFb Fba             '<== Crt Fba
                  ExpPj Pj              '<== Exp
                  OpnFb Acs, Fba
     Set DistPj = PjzAcs(Acs)
                  AddRfzDistPj DistPj   '<== Add Rf
                  LoadBas DistPj        '<== Load Bas
Dim P$:       P = Srcp(Pj)
Dim Frm$(): Frm = FrmFfny(P)
Dim F: For Each FrmFfn In Itr(Frm)
    Dim N$: N = RmvExt(RmvExt(F))
    Acs.LoadFromText acForm, N, F       '<== Load Frm
Next
QuitAcs Acs
End Sub

