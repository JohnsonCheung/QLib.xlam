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
Dim A As Drs: A = DMthzP(P)
Dim B As Drs: B = DrswDup(A, "Mthn")
Dim C As Drs: C = DrseDup(B, "Mthn Mdn") '<==
Dim D As Drs: D = AddColzMthL(C)
Dim E As Drs: E = AddColzValIdzCntzDrs(D, "MthL")
DDupMthzP = SrtDrs(E)
End Function


Function FmtDupMthWs(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetLcWdt Lo, "MthL", 10
SetLcWrp Lo, "MthL", False
End Function


Private Function AddColzMthL(Mthn As Drs) As Drs
Dim A():  A = DrszSel(Mthn, "Md Mthn Ty").Dry
Dim B$(): B = MthLAyzDry_Md_Mthn_ShtMthTy(A)
Dim C As Drs: C = DrsAddColzNmVy(Mthn, "MthL", B)
AddColzMthL = C
End Function


Function DupMthWszP(P As VBProject) As Worksheet
Set DupMthWszP = FmtDupMthWs(WszDrs(DDupMthzP(P), "DupMth"))
End Function


Private Function MthLAyzDry_Md_Mthn_ShtMthTy(Dry()) As String()
Dim Dr, M As CodeModule, Mthn, ShtMthTy$
For Each Dr In Itr(Dry)
    Set M = Md(Dr(0))
    Mthn = Dr(1)
    ShtMthTy = Dr(2)
    PushI MthLAyzDry_Md_Mthn_ShtMthTy, MthLzMTN(M, ShtMthTy, Mthn)
Next
End Function

Private Sub Z_DDupMthP()
BrwDrs DDupMthP
End Sub


Private Sub Z()
'Z_PjDupMthNyWithLinesId
MIde_Mth_Dup:
End Sub

