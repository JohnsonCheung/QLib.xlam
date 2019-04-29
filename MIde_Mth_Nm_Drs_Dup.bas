Attribute VB_Name = "MIde_Mth_Nm_Drs_Dup"
Option Explicit
Function DupMthWsInPj(Optional Vis As Boolean) As Worksheet
Set DupMthWsInPj = DupMthWszPj(CurPj)
End Function

Function DupMthWszPj(A As VBProject, Optional Vis As Boolean) As Worksheet
Set DupMthWszPj = DupMthWsFmt(WszDrs(DupMthDrszPj(A), "DupMth", Vis))
End Function

Function DupMthWsFmt(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetLcWdt Lo, "MthLines", 10
SetLcWrp Lo, "MthLines", False
End Function

Private Sub Z_DupMthDrsInPj()
B DupMthDrsInPj
End Sub
Function DupMthDrsInPj() As Drs
Set DupMthDrsInPj = DupMthDrszPj(CurPj)
End Function

Private Function DupMthDrszPj(A As VBProject) As Drs
Dim B As Drs: Set B = MthNmDrszPj(A, "-Mod -Pub")
Dim C As Drs: Set C = DrswDup(B, "MthNm")
Dim D As Drs: Set D = DrseDup(C, "MthNm Md") '<==
Dim E As Drs: Set E = AddColzMthLines(D)
Dim F As Drs: Set F = AddColzValIdzCntzDrs(E, "MthLines")

Set DupMthDrszPj = DrsSrt(F)
End Function

Private Function AddColzMthLines(MthNmDrs As Drs) As Drs
Dim A():  A = DrsSel(MthNmDrs, "Md MthNm Ty").Dry
Dim B$(): B = MthLinesAyzDry_Md_MthNm_ShtMthTy(A)
Dim C As Drs: Set C = DrsAddColzNmVy(MthNmDrs, "MthLines", B)
Set AddColzMthLines = C
End Function

Private Function MthLinesAyzDry_Md_MthNm_ShtMthTy(Dry()) As String()
Dim Dr, M As CodeModule, MthNm$, ShtMthTy$
For Each Dr In Itr(Dry)
    Set M = Md(Dr(0))
    MthNm = Dr(1)
    ShtMthTy = Dr(2)
    PushI MthLinesAyzDry_Md_MthNm_ShtMthTy, MthLinesByMdNmTy(M, MthNm, ShtMthTy, WiTopRmk:=True)
Next
End Function
