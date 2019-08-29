Attribute VB_Name = "QIde_F_DupMth"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Dup_X."
Private Const Asm$ = "QIde"

Function DupMthWsP() As Worksheet
Set DupMthWsP = DupMthWszP(CPj)
End Function

Function DoDupMthP() As Drs
DoDupMthP = DoDupMthzP(CPj)
End Function

Private Function DoDupMthzP(P As VBProject) As Drs
Dim A As Drs: A = DoMthzP(P)
Dim B As Drs: B = DwDup(A, "Mthn")
Dim C As Drs: C = DeDupzFF(B, "Mthn Mdn") '<==
Dim D As Drs: D = AddColzMthL(C)
Dim E As Drs: E = AddColzValIdqCnt(D, "MthL")
DoDupMthzP = SrtDrs(E)
End Function

Function FmtDupMthWs(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetWdtLc Lo, "MthL", 10
SetWrpLc Lo, "MthL", False
End Function

Private Function AddColzMthL(Mthn As Drs) As Drs
Dim A():  A = SelDrs(Mthn, "Mdn Mthn Ty").Dy
Dim B$(): B = MthLAyzDy_Mdn_Mthn_ShtMthTy(A)
Dim C As Drs: C = AddColzVy(Mthn, "MthL", B)
AddColzMthL = C
End Function

Function DupMthWszP(P As VBProject) As Worksheet
Set DupMthWszP = FmtDupMthWs(WszDrs(DoDupMthzP(P), "DupMth"))
End Function

Private Function MthLAyzDy_Mdn_Mthn_ShtMthTy(Dy()) As String()
Dim Dr, M As CodeModule, Mthn, ShtMthTy$
For Each Dr In Itr(Dy)
    Set M = Md(Dr(0))
    Mthn = Dr(1)
    ShtMthTy = Dr(2)
    PushI MthLAyzDy_Mdn_Mthn_ShtMthTy, MthLzNmTy(M, Mthn, ShtMthTy)
Next
End Function

Private Sub Z_DoDupMthP()
BrwDrs DoDupMthP
End Sub

Private Sub Z()
'Z_PjDupMthNyWithLinesId
MIde_Mth_Dup:
End Sub


'
