Attribute VB_Name = "MIde_Ens_OptLin"
Option Explicit
Const CMod$ = "MIde_Ens_Option."
Const OptLinoExplicit$ = "Option Explicit"
Const OptLinoCmpBin$ = "Option Compare Binary"
Const OptLinoCmpDb$ = "Option Compare Database"
Const OptLinoCmpTxt$ = "Option Compare Text"

Sub EnsOptLinPj()
EnsOptLinzPj CurPj
End Sub

Sub EnsMdOptLin()
EnsOptLinzMd CurMd
End Sub

Private Sub EnsOptLinzPj(Pj As VBProject)
Dim C As VBComponent
For Each C In Pj.VBComponents
    EnsOptLinzMd C.CodeModule
Next
End Sub

Private Sub EnsOptLinzMd(A As CodeModule)
RmvOptLin A, OptLinoCmpDb
RmvOptLin A, OptLinoCmpBin
RmvOptLin A, OptLinoCmpTxt
EnsOptLin A, OptLinoExplicit
'EnsCLib A
End Sub
Private Sub EnsCLibzPj(A As VBProject, Optional B As eLibNmTy)
If A.Protection = vbext_pp_locked Then Exit Sub
Dim C As VBComponent
For Each C In A.VBComponents
    EnsCLib C.CodeModule, B
Next
End Sub

Function LnozAftOpt%(A As CodeModule)
Dim OptFnd As Boolean, J%, IsOpt As Boolean
For J = 1 To A.CountOfDeclarationLines
    IsOpt = IsOptLin(A.Lines(J, 1))
    Select Case True
    Case OptFnd And IsOpt:
    Case OptFnd: LnozAftOpt = J: Exit Function
    Case IsOpt: OptFnd = True
    End Select
Next
LnozAftOpt = J
End Function

Private Function OptLno%(A As CodeModule, OptLin$)
Dim J&
For J = 1 To A.CountOfDeclarationLines
   If A.Lines(J, 1) = OptLin Then OptLno = J: Exit Function
Next
End Function

Private Sub EnsOptLin(A As CodeModule, OptLin$)
Const CSub$ = CMod & "EnsOptLin"
If A.CountOfLines = 0 Then Exit Sub
If OptLno(A, OptLin) > 0 Then Exit Sub
A.InsertLines 1, OptLin
InfoLin CSub, "[" & OptLin & "] is Inserted", "Md", MdNm(A)
End Sub

Private Sub RmvOptLin(A As CodeModule, OptLin$)
Const CSub$ = CMod & "RmvOptLin"
Dim I%: I = OptLno(A, OptLin)
If I = 0 Then Exit Sub
A.DeleteLines I
Info CSub, "[" & OptLin & "] line is deleted", "Md Lno", MdNm(A), I
End Sub
