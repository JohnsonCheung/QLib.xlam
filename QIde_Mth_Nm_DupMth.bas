Attribute VB_Name = "QIde_Mth_Nm_DupMth"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Dup_X."
Private Const Asm$ = "QIde"
Function DupMthWsP() As Worksheet
Set DupMthWsP = DupMthWszP(CPj)
End Function


Function Drs_DupMthP() As Drs
Drs_DupMthP = Drs_DupMthzP(CPj)
End Function


Private Function Drs_DupMthzP(P As VBProject) As Drs
Dim B As Drs: B = Drs_MthnzP(P)
Dim C As Drs: C = DrswDup(B, "Mthn")
Dim D As Drs: D = DrseDup(C, "Mthn Md") '<==
Dim E As Drs: E = AddColzMthLines(D)
Dim F As Drs: F = AddColzValIdzCntzDrs(E, "MthLines")
Drs_DupMthzP = SrtDrs(F)
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
Set DupMthWszP = FmtDupMthWs(WszDrs(Drs_DupMthzP(P), "DupMth"))
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

Private Sub Z_Drs_DupMthP()
BrwDrs Drs_DupMthP
End Sub





Function SamMthLinesMthDnDry(QMthnLDrs As Drs, Vbe As Vbe) As Variant()
Dim Gp(): 'Gp = DupQMthny_Blk(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupQMthnyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
'SamMthLinesMthDnDry = O
End Function

Private Function IfShwNoDupMsg(MthDNy$(), Mthn) As Boolean
IfShwNoDupMsg = False
Select Case Si(MthDNy)
Case 0: Inf CSub, "No such method in CVbe", "Mthn", Mthn
Case 1: Inf CSub, "No dup method", "MthDn", MthDNy(0)
Case Else: IfShwNoDupMsg = True
End Select
End Function

Function DupQMthnyGp_IsDup(Ny) As Boolean
'DupQMthnyGp_IsDup = IsEqzAllEle(MapAy(Ny, "FunFNm_MthLines"))
End Function

Function DupQMthnyGp_IsVdt(DupQMthnyGp$()) As Boolean
Dim A$(): A = DupQMthnyGp
If Si(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupQMthnyGp_IsVdt = True
End Function

Function DupQMthnyBlkAllSameCnt%(A)
If Si(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupQMthnyGp_IsDup(Gp) Then O = O + 1
Next
DupQMthnyBlkAllSameCnt = O
End Function

Function DupQDr_MthnsP() As Drs
DupQDr_MthnsP = DupQDr_MthnszP(CPj)
End Function

Function DupQDr_MthnszP(P As VBProject) As Drs
'DupQDr_MthnszP = DrszFF("Pj Md Mth Ty Mdy", DupQDry_MthnzP(A))
End Function

Function DupQDry_MthnPj() As Variant()
DupQDry_MthnPj = DupQDry_MthnzP(CPj)
End Function

Function DupQDry_MthnzP(P As VBProject) As Variant()
Dim Dry(), Dry1(), Dry2()
'Dry = QDry_MthnzP(A, "-Mod") ' Pjn Mdn Mthn Ty Mdy
'Dry1 = DryeCEv(Dry, 4, "Prv")
Dry2 = DrywDupCC(Dry, LngAp(2))
DupQDry_MthnzP = SrtDryzCol(Dry2, 2)
End Function

Function DupIxyzDry(Dry(), CCIxy&()) As Long()

End Function

Function DupQDry_MthnVbe() As Variant()
DupQDry_MthnVbe = DupQDry_MthnzV(CVbe)
End Function

Function DupQDry_MthnzQMthny(QMthny$()) As Variant()
DupQDry_MthnzQMthny = DrywDupCC(DryzDotAy(QMthny), LngAp(2))
End Function

Function DupQDry_MthnzV(A As Vbe) As Variant()
DupQDry_MthnzV = DupQDry_MthnzQMthny(QMthnyzV(A))
End Function

Private Sub ZZ()
'Z_PjDupMthnyWithLinesId
MIde_Mth_Dup:
End Sub

