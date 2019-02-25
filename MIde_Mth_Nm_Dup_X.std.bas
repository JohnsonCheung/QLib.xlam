Attribute VB_Name = "MIde_Mth_Nm_Dup_X"
Option Explicit

Function SamMthLinesMthQDNmDry(MthQNmLDrs As Drs, Vbe As Vbe) As Variant()
Dim Gp(): 'Gp = DupMthQDNy_GpAy(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupMthQDNyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
'SamMthLinesMthQDNmDry = O
End Function

Private Function IfShwNoDupMsg(MthQDNy$(), MthNm) As Boolean
IfShwNoDupMsg = False
Select Case Sz(MthQDNy)
Case 0: Info CSub, "No such method in CurVbe", "MthNm", MthNm
Case 1: Info CSub, "No dup method", "MthQDNm", MthQDNy(0)
Case Else: IfShwNoDupMsg = True
End Select
End Function

Function DupMthQDNyGp_IsDup(Ny) As Boolean
'DupMthQDNyGp_IsDup = IsAllEleEqAy(AyMap(Ny, "FunFNm_MthLines"))
End Function

Function DupMthQDNyGp_IsVdt(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupMthQDNyGp_IsVdt = True
End Function

Function DupMthQDNyGpAyAllSameCnt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupMthQDNyGp_IsDup(Gp) Then O = O + 1
Next
DupMthQDNyGpAyAllSameCnt = O
End Function

Function DupPjLinesIdMthNy(A As VBProject) As String()
Dim Dic As New Dictionary, N
Dim M
'For Each N In Itr(PjDupMthNy(A))
'    PushI DupPjLinesIdMthNy, N & "." & X1(A, N, Dic)
'Next
End Function

Function DupMthQDNmDryPj(A As VBProject) As String()
Dim Dry()
'Dry = MthQDNmDryPj(A) ' PjNm MdNm MthNm
Dry = SrtDry(Dry, 2)
Dry = DrySelIxAp(Dry, 1, 2) ' MthNm MdNm
'PjDupMthNy = DotNyDry(Dry)
End Function
Function DupMthQDNmDry() As Variant()
DupMthQDNmDry = DupMthQDNmDryzbe(CurVbe)
End Function
Function MthQDNmDryzbe(A As Vbe) As Variant()
Dim P As VBProject
For Each P In A.VBProjects
'    PushIAy MthQDNmDryzbe, MthQDNmDryPj(P)
Next
End Function
Function DupMthQDNmDryMthQDNmDry(A()) As Variant()

End Function
Function DupMthQDNmDryzbe(A As Vbe) As Variant()
DupMthQDNmDryzbe = DupMthQDNmDryMthQDNmDry(MthQDNmDryzbe(A))
End Function

Private Sub Z()
'Z_PjDupMthNyWithLinesId
MIde_Mth_Dup:
End Sub

