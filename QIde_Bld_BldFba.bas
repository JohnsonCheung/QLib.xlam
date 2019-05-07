Attribute VB_Name = "QIde_Bld_BldFba"
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Fba."
Private Const Asm$ = "QIde"
Private Type A
    Pj As VBProject
                            End Type
Private Type R1
    Srcp   As String
    DistPj As VBProject
                            End Type
Private Type R2
    Acs    As Access.Application
    DistPj As VBProject
                            End Type
Private Type R5
    FrmFfnSy() As String
                            End Type
Private A As A
Private R1 As R1
Private R2 As R2
Private R5 As R5
Sub BldFbaP()
BldFbazPj CurPj
End Sub
Private Sub Nop(): End Sub
Sub BldFbazPj(Pj As VBProject)
Set A.Pj = Pj
R1_ExpPj: Nop
    ExpPj A.Pj
    R1.Srcp = Srcp(A.Pj)
    R1.DistFba = DistFba(R1.Srcp)
R2_CrtFba: NoOp
    DltFfnIf R1.DistFba
    CrtFb R1.DistFba
    Set R2.Acs = New Access.Application
    OpnFb R2.Acs, R1.DistPj
    Set R2.DistPj = PjzAcs(R2.Acs)
R3_AddRf: NoOp
    AddRfzDistPj R2.DistPj
R4_LoadBas: Nop
    LoadBas R2.DistPj
R5_LoadFrm: Nop
    R5.FrmFfnSy = FrmFfnSy(R1.Srcp)
    Dim FrmFfn$, I, N$
    For Each I In Itr(R5.FrmFfnSy)
        FrmFfn = I
        N$ = RmvExt(RmvExt(FrmFfn))
        R2.Acs.LoadFromText acForm, RmvExt(RmvExt(FrmFfn)), FrmFfn
    Next
R6_QuitAcs: Nop
    QuitAcs R2.Acs
End Sub

