Attribute VB_Name = "MIde_Mth_Nm_Dup_X"
Option Explicit
Function SamMthLinesMthDNmDry(MthQNmLDrs As Drs, Vbe As Vbe) As Variant()
Dim Gp(): 'Gp = DupMthQNy_GpAy(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupMthQNyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
'SamMthLinesMthDNmDry = O
End Function

Private Function IfShwNoDupMsg(MthDNy$(), MthNm) As Boolean
IfShwNoDupMsg = False
Select Case Si(MthDNy)
Case 0: Inf CSub, "No such method in CurVbe", "MthNm", MthNm
Case 1: Inf CSub, "No dup method", "MthDNm", MthDNy(0)
Case Else: IfShwNoDupMsg = True
End Select
End Function

Function DupMthQNyGp_IsDup(Ny) As Boolean
'DupMthQNyGp_IsDup = IsAllEleEq(AyMap(Ny, "FunFNm_MthLines"))
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

Function DupMthQNyGpAyAllSameCnt%(A)
If Si(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupMthQNyGp_IsDup(Gp) Then O = O + 1
Next
DupMthQNyGpAyAllSameCnt = O
End Function

Function DupMthQNmDrsInPj() As Drs
DupMthQNmDrsInPj = DupMthQNmDrszPj(CurPj)
End Function

Function DupMthQNmDrszPj(A As VBProject) As Drs
DupMthQNmDrszPj = DrszFF("Pj Md Mth Ty Mdy", DupMthQNmDryzPj(A))
End Function

Function DupMthQNmDryPj() As Variant()
DupMthQNmDryPj = DupMthQNmDryzPj(CurPj)
End Function

Function DupMthQNmDryzPj(A As VBProject) As Variant()
Dim Dry(), Dry1(), Dry2()
Dry = MthQNmDryzPj(A, "-Mod") ' PjNm MdNm MthNm Ty Mdy
'Dry1 = DryeCEv(Dry, 4, "Prv")
Dry2 = DrywDupCC(Dry, LngAy(2))
DupMthQNmDryzPj = DrySrtzCol(Dry2, 2)
End Function

Function DupIxAyzDry(Dry(), CCIxAy&()) As Long()

End Function

Function DupMthQNmDryVbe() As Variant()
DupMthQNmDryVbe = DupMthQNmDryzVbe(CurVbe)
End Function

Function DupMthQNmDryzMthQNy(MthQNy$()) As Variant()
DupMthQNmDryzMthQNy = DrywDupCC(DryzDotAy(MthQNy), LngAy(2))
End Function

Function DupMthQNmDryzVbe(A As Vbe) As Variant()
DupMthQNmDryzVbe = DupMthQNmDryzMthQNy(MthQNyzVbe(A, "-Mod -Pub"))
End Function

Private Sub Z()
'Z_PjDupMthNyWithLinesId
MIde_Mth_Dup:
End Sub

