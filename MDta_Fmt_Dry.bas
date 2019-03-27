Attribute VB_Name = "MDta_Fmt_Dry"
Option Explicit
Private Sub A_Main()
FmtDryAsSpcSep:
FmtDry:
End Sub
Sub BrwDry(A(), Optional MaxColWdt% = 100, Optional BrkCC, Optional ShwZer As Boolean)
BrwAy FmtDry(A, MaxColWdt, BrkCC, ShwZer)
End Sub
Sub BrwDryzSpc(A(), Optional MaxColWdt% = 100, Optional ShwZer As Boolean)
BrwAy FmtDryAsSpcSep(A, MaxColWdt, ShwZer)
End Sub

Function FmtDry(Dry(), _
Optional MaxColWdt% = 100, _
Optional BrkCC, _
Optional ShwZer As Boolean) _
As String()
If Si(Dry) = 0 Then Exit Function
Dim Dry1(): Dry1 = DryStrfy(Dry, ShwZer, MaxColWdt)
Dim W%(): W = WdtAyzDry(Dry1)
Dim Dry2(): Dry2 = DryAlignColzWdt(Dry, W)
Dim Sep$(): Sep = SepDr(W)
Dim Dry3(): Dry3 = DryInsBrk(Dry2, BrkCC, Sep)
Dim SepLin$: SepLin = "|" & Jn(Sep, "|") & "|"
FmtDry = SyAddSorSyAp(EmpSy, SepLin, FmtDryByJnCell(Dry3), SepLin)
End Function

Sub DmpDryAsSpcSep(Dry())
D FmtDryAsSpcSep(Dry)
End Sub
Sub DmpDry(Dry())
D FmtDry(Dry)
End Sub

Function FmtDryAsSpcSep(Dry(), _
Optional MaxColWdt% = 100, _
Optional ShwZer As Boolean) As String()
If Si(Dry) = 0 Then Exit Function
Dim Dr
For Each Dr In DryFmtCommon(Dry, MaxColWdt, ShwZer)
    PushI FmtDryAsSpcSep, JnSpc(Dr)
Next
End Function


