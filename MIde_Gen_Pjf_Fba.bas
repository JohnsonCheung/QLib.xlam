Attribute VB_Name = "MIde_Gen_Pjf_Fba"
Option Explicit
Private Type A
    Acs As Access.Application
    Scrp As String
End Type
Private Type B
    DistPj As VBProject
End Type
Private A As A, B As B
Sub GenFba(Scrp, Acs As Access.Application)
ThwNotSrcp Scrp
Set A.Acs = Acs
A.Scrp = Scrp
B = Gen_CrtFba
Gen_AddRf
Gen_LoadBas
Gen_LoadFrm
Gen_Cls
End Sub
Private Function Gen_CrtFba() As B
Dim Fba$: Fba = DistFba(A.Scrp)
DltFfnIf Fba
CrtFb Fba
OpnFb A.Acs, Fba
Set Gen_CrtFba.DistPj = PjzAcs(A.Acs)
End Function
Private Sub Gen_Cls()
A.Acs.CloseCurrentDatabase
End Sub
Private Sub Gen_AddRf()
AddRfzDistPj B.DistPj
End Sub
Private Sub Gen_LoadBas()
LoadBas B.DistPj
End Sub
Private Function LoadFrm_FrmFfnAy() As String()
Dim Ffn
For Each Ffn In FfnSy(A.Scrp)
    If LoadFrm_IsFrmFfn(Ffn$) Then
        PushI LoadFrm_FrmFfnAy, Ffn
    End If
Next
End Function

Private Function LoadFrm_IsFrmFfn(Ffn$) As Boolean
LoadFrm_IsFrmFfn = HasSfx(Ffn$, ".frm.txt")
End Function
Sub Gen_LoadFrm()
Dim FrmFfn, N$
For Each FrmFfn In Itr(LoadFrm_FrmFfnAy)
    N$ = RmvExt(RmvExt(FrmFfn))
    A.Acs.LoadFromText acForm, RmvExt(RmvExt(FrmFfn)), FrmFfn
Next
End Sub
