Attribute VB_Name = "QIde_Mth_LinAy_Drs"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Liny_Drs."
Function MthLinAyP() As String()
MthLinAyP = StrCol(DMthP, "MthLin")
End Function
Function MthLinAyzP(P As VBProject) As String()
MthLinAyzP = StrCol(DMthzP(P), "MthLin")
End Function

Function MthLinAyV() As String()
MthLinAyV = MthLinAyzV(CVbe)
End Function

Function MthLinAyzV(V As Vbe) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinAyzV, MthLinAyzP(P)
Next
End Function

