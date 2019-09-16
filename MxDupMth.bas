Attribute VB_Name = "MxDupMth"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDupMth."
Public Const FFoDupMth$ = FFoMthc

Function DupMthWsP() As Worksheet
Set DupMthWsP = DupMthWszP(CPj)
End Function

Function DoDupMthP(Optional InclPrv As Boolean, Optional IsExactDup As Boolean) As Drs
DoDupMthP = DoDupMthzP(CPj, InclPrv, IsExactDup)
End Function

Private Function DoDupMthzP(P As VBProject, Optional InclPrv As Boolean, Optional IsExactDup As Boolean) As Drs
Dim A As Drs: A = DwEq(DoMthczP(P), "MdTy", "Std")
Dim A1 As Drs:
    If InclPrv Then
        A1 = A
    Else
        A1 = DwNe(A, "Mdy", "Prv")
    End If
Dim B As Drs: B = DwDup(A1, "Mthn")
Dim C As Drs: C = SrtDrs(B, "Mthn")
Dim D As Drs: D = AddColzValIdqCnt(C, "Mthl")
If IsExactDup Then
    DoDupMthzP = DwDup(D, "MthlId")
Else
    DoDupMthzP = D
End If
End Function

Function FmtDupMthWs(DupMthWs As Worksheet) As Worksheet
Dim Lo As ListObject: Set Lo = FstLo(DupMthWs)
SetLcWdt Lo, "MthL", 10
SetLcWrp Lo, "MthL", False
End Function

Function DupMthWszP(P As VBProject) As Worksheet
Set DupMthWszP = FmtDupMthWs(WszDrs(DoDupMthzP(P), "DupMth"))
End Function

Private Sub Z_DoDupMthP()
BrwDrs DoDupMthP
End Sub

Private Sub Z()
'Z_PjDupMthNyWithLinesId
MIde_Mth_Dup:
End Sub