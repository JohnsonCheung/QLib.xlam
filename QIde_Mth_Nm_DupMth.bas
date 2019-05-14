Attribute VB_Name = "QIde_Mth_Nm_DupMth"
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Dup_X."
Private Const Asm$ = "QIde"
Function DupMthWsP(Optional Vis As Boolean) As Worksheet
Set DupMthWsP = DupMthWszP(CPj)
End Function


Function DrsOfDupMthP() As Drs
DrsOfDupMthP = DrsOfDupMthzP(CPj)
End Function


Private Function DrsOfDupMthzP(P As VBProject) As Drs
Dim B As Drs: B = DrsOfMthnzP(P, "-Mod -Pub")
Dim C As Drs: C = DrswDup(B, "Mthn")
Dim D As Drs: D = DrseDup(C, "Mthn Md") '<==
Dim E As Drs: E = AddColzMthLines(D)
Dim F As Drs: F = AddColzValIdzCntzDrs(E, "MthLines")
DrsOfDupMthzP = DrsSrt(F)
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
Set DupMthWszP = FmtDupMthWs(WszDrs(DrsOfDupMthzP(P), "DupMth"))
End Function



Private Function MthLinesAyzDry_Md_Mthn_ShtMthTy(Dry()) As String()
Dim Dr, M As CodeModule, Mthn, ShtMthTy$
For Each Dr In Itr(Dry)
    Set M = Md(Dr(0))
    Mthn = Dr(1)
    ShtMthTy = Dr(2)
    PushI MthLinesAyzDry_Md_Mthn_ShtMthTy, MthLineszMTN(M, ShtMthTy, Mthn, WiTopRmk:=True)
Next
End Function

Private Sub Z_DrsOfDupMthP()
BrwDrs DrsOfDupMthP
End Sub





Function SamMthLinesMthDnDry(MthQNmLDrs As Drs, Vbe As Vbe) As Variant()
Dim Gp(): 'Gp = DupMthQNy_Blk(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupMthQNyGp_IsDup(Ny) Then
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

Function DupMthQNyGp_IsDup(Ny) As Boolean
'DupMthQNyGp_IsDup = IsEqzAllEle(MapAy(Ny, "FunFNm_MthLines"))
End Function

Function DupMthQNyGp_IsVdt(DupMthQNyGp$()) As Boolean
Dim A$(): A = DupMthQNyGp
If Si(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupMthQNyGp_IsVdt = True
End Function

Function DupMthQNyBlkAllSameCnt%(A)
If Si(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupMthQNyGp_IsDup(Gp) Then O = O + 1
Next
DupMthQNyBlkAllSameCnt = O
End Function

Function DupMthQNmDrsP() As Drs
DupMthQNmDrsP = DupMthQNmDrszP(CPj)
End Function

Function DupMthQNmDrszP(P As VBProject) As Drs
'DupMthQNmDrszP = DrszFF("Pj Md Mth Ty Mdy", DupMthQNmDryzP(A))
End Function

Function DupMthQNmDryPj() As Variant()
DupMthQNmDryPj = DupMthQNmDryzP(CPj)
End Function

Function DupMthQNmDryzP(P As VBProject) As Variant()
Dim Dry(), Dry1(), Dry2()
'Dry = MthQNmDryzP(A, "-Mod") ' Pjn Mdn Mthn Ty Mdy
'Dry1 = DryeCEv(Dry, 4, "Prv")
Dry2 = DrywDupCC(Dry, Lngy(2))
DupMthQNmDryzP = DrySrtzCol(Dry2, 2)
End Function

Function DupIxyzDry(Dry(), CCIxy&()) As Long()

End Function

Function DupMthQNmDryVbe() As Variant()
DupMthQNmDryVbe = DupMthQNmDryzV(CVbe)
End Function

Function DupMthQNmDryzMthQNy(MthQNy$()) As Variant()
DupMthQNmDryzMthQNy = DrywDupCC(DryzDotAy(MthQNy), Lngy(2))
End Function

Function DupMthQNmDryzV(A As Vbe) As Variant()
DupMthQNmDryzV = DupMthQNmDryzMthQNy(MthQNyzV(A, "-Mod -Pub"))
End Function

Private Sub ZZ()
'Z_PjDupMthnyWithLinesId
MIde_Mth_Dup:
End Sub

