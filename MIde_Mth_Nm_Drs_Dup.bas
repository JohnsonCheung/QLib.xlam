Attribute VB_Name = "MIde_Mth_Nm_Drs_Dup"
Option Explicit

Function DupMthDRsPj() As DRs
Set DupMthDRsPj = DupMthDrszPj(CurPj)
End Function

Private Function DupMthDrszPj(A As VBProject) As DRs
Dim B As DRs: Set B = MthNmDrszPj(A, "-Mod -Pub")
Dim C As DRs: Set C = DrswDup(B, "MthNm")
Dim D As DRs: Set D = DrseDup(C, "MthNm Md") '<==
Dim E As DRs: Set E = AddColzMthLines(D)
Dim F As DRs: Set F = AddColzValIdzCntzDrs(E, "MthLines")
Set DupMthDrszPj = F
End Function


Private Function AddColzMthLines(MthNmDrs As DRs) As DRs
Stop
Dim A():  A = DrsSel(MthNmDrs, "Md MthNm Ty").Dry
Dim B$(): B = MthLinesAyzDry_Md_MthNm_ShtMthTy(A)
Dim C As DRs: Set C = AddColzColVyDrs(MthNmDrs, "MthLines", B)
Set AddColzMthLines = C
End Function

Private Function MthLinesAyzDry_Md_MthNm_ShtMthTy(Dry()) As String()
Dim Dr, M As CodeModule, MthNm$, ShtMthTy$
For Each Dr In Itr(Dry)
    Set M = Md(Dr(0))
    MthNm = Dr(1)
    ShtMthTy = Dr(2)
    PushI MthLinesAyzDry_Md_MthNm_ShtMthTy, MthLineszMdNmTy(M, MthNm, ShtMthTy, WithTopRmk:=True)
Next
End Function
