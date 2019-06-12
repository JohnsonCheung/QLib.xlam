Attribute VB_Name = "QIde_F_Mth_DupMth"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Dup_X."
Private Const Asm$ = "QIde"
Function DupMthWsP() As Worksheet
Set DupMthWsP = DupMthWszP(CPj)
End Function


Function DDupMthP() As Drs
DDupMthP = DDupMthzP(CPj)
End Function


Private Function DDupMthzP(P As VBProject) As Drs
Stop
Dim B As Drs: B = DMthnzP(P)
Dim C As Drs: C = DrswDup(B, "Mthn")
Dim D As Drs: D = DrseDup(C, "Mthn Md") '<==
Dim E As Drs: E = AddColzMthLines(D)
Dim F As Drs: F = AddColzValIdzCntzDrs(E, "MthLines")
DDupMthzP = SrtDrs(F)
End Function


Function FmtDupMthWs(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetLcWdt Lo, "MthLines", 10
SetLcWrp Lo, "MthLines", False
End Function


Private Function AddColzMthLines(Mthn As Drs) As Drs
Dim A():  A = SelDrs(Mthn, "Md Mthn Ty").Dry
Dim B$(): B = MthLinesAyzDry_Md_Mthn_ShtMthTy(A)
Dim C As Drs: C = DrsAddColzNmVy(Mthn, "MthLines", B)
AddColzMthLines = C
End Function


Function DupMthWszP(P As VBProject) As Worksheet
Set DupMthWszP = FmtDupMthWs(WszDrs(DDupMthzP(P), "DupMth"))
End Function


Private Function MthLinesAyzDry_Md_Mthn_ShtMthTy(Dry()) As String()
Dim Dr, M As CodeModule, Mthn, ShtMthTy$
For Each Dr In Itr(Dry)
    Set M = Md(Dr(0))
    Mthn = Dr(1)
    ShtMthTy = Dr(2)
    PushI MthLinesAyzDry_Md_Mthn_ShtMthTy, MthLineszMTN(M, ShtMthTy, Mthn)
Next
End Function

Private Sub Z_DDupMthP()
BrwDrs DDupMthP
End Sub


Private Sub ZZ()
'Z_PjDupMthnyWithLinesId
MIde_Mth_Dup:
End Sub

