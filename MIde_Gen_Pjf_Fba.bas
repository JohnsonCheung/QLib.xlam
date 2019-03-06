Attribute VB_Name = "MIde_Gen_Pjf_Fba"
Option Explicit

Sub GenFba(SrcPthInst, Optional Acs As Access.Application)
ThwNotSrcPthInst SrcPthInst
Dim A As Access.Application: Set A = DftAcs(A)
Dim Fba$: Fba = DistFba(SrcPthInst)
CrtFb Fba
OpnFb A, Fba
Dim Pj As VBProject: Set Pj = A.Vbe.ActiveVBProject
AddRfzPj Pj
LoadBas Pj
LoadFrm Pj
ClsDbzAcs A
CpyFilzToPth Fba, AddFdrEns(ParPth(ParPth(Pth(Fba))), "Dist"), OvrWrt:=True
If IsNothing(A) Then AcsQuit A
End Sub
Private Sub LoadFrm(A As VBProject)
Stop
End Sub
Private Sub LoadFrmzAcs(A As Access.Application, SrcPth)
Dim FrmFfn, N$
For Each FrmFfn In Itr(FrmFfnAy(SrcPth))
    N$ = RmvExt(RmvExt(FrmFfn))
    A.LoadFromText acForm, RmvExt(RmvExt(FrmFfn)), FrmFfn
Next
End Sub

Private Function FrmFfnAy(SrcPth) As String()
Dim Ffn
For Each Ffn In FfnAy(SrcPth)
    If IsFrmFfn(Ffn) Then
        PushI FrmFfnAy, Ffn
    End If
Next
End Function
Private Function IsFrmFfn(Ffn) As Boolean
IsFrmFfn = HasSfx(Ffn, ".frm.txt")
End Function
