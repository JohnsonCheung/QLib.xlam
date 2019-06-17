Attribute VB_Name = "QIde_Vbe"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Vbe."
Private Const Asm$ = "QIde"

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function
Sub DmpIsPjSav()
DmpDrs DrszIsPjSav(CVbe)
End Sub
Function Dry_IsPjSav(A As Vbe) As Variant()
Dim I As VBProject
For Each I In A.VBProjects
    PushI Dry_IsPjSav, Array(I.Saved, I.Name, I.GenFileName)
Next
End Function
Function DrszIsPjSav(A As Vbe) As Drs
DrszIsPjSav = DrszFF("IsSav Pjn GenFfn", Dry_IsPjSav(A))
End Function

Function PjzV(A As Vbe, Pjn$) As VBProject
Set PjzV = A.VBProjects(Pjn)
End Function

Function PjzPjf(Vbe As Vbe, Pjf) As VBProject
Dim I As VBProject
For Each I In Vbe.VBProjects
    If PjfzP(I) = Pjf Then Set PjzPjf = I: Exit Function
Next
End Function

Function MdDryzV(A As Vbe) As Variant()
Dim C, Pnm$, P As VBProject
For Each P In A.VBProjects
    Push MdDryzV, MdDr(C.CodeModule)
Next
End Function
Function MdDr(M As CodeModule) As Variant()

End Function

Sub SavVbe(A As Vbe)
Dim P As VBProject
For Each P In A.VBProjects
    SavPj P
Next
End Sub

Property Get PjfyV() As String()
PjfyV = PjfyzV(CVbe)
End Property

Function PjfyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNonBlank PjfyzV, Pjf(P)
Next
End Function

Function PjnyV() As String()
PjnyV = PjnyzV(CVbe)
End Function

Function PjnyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushI PjnyzV, P.Name
Next
End Function

Function SrtRptV() As String()
SrtRptV = SrtRptzV(CVbe)
End Function

Function HasBarzV(A As Vbe, BarNm) As Boolean
HasBarzV = HasItn(A.CommandBars, BarNm)
End Function

Function HasPj(A As Vbe, Pjn$) As Boolean
HasPj = HasItn(A.VBProjects, Pjn)
End Function

Function HasPjfzV(A As Vbe, Pjf) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If PjfzP(P) = Pjf Then HasPjfzV = True: Exit Function
Next
End Function

Function SrtRptzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy SrtRptzV, SrtRptzP(P)
Next
End Function

Private Sub Z_VbeFunPfx()
'D Vbe_MthPfx(CVbe)
End Sub

Private Sub Z_MthNyzV()
Brw MthNyzV(CVbe)
End Sub


Private Sub ZZ()
Dim A
Dim B As Vbe
Dim C$
Dim D As Boolean
Dim XX
CvVbe A
End Sub

