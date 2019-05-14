Attribute VB_Name = "QIde_Vbe"
Option Explicit
Private Const CMod$ = "MIde_Vbe."
Private Const Asm$ = "QIde"

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function
Sub DmpIsPjSav()
DmpDrs DrszIsPjSav(CVbe)
End Sub
Function DryOfIsPjSav(A As Vbe) As Variant()
Dim I As VBProject
For Each I In A.VBProjects
    PushI DryOfIsPjSav, Array(I.Saved, I.Name, I.GenFileName)
Next
End Function
Function DrszIsPjSav(A As Vbe) As Drs
DrszIsPjSav = DrszFF("IsSav Pjn GenFfn", DryOfIsPjSav(A))
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

Function MdDryzV(A As Vbe, Optional WhStr$) As Variant()
Dim P, C, Pnm$, Pj As VBProject
For Each P In PjItr(A, WhStr)
    Set Pj = P
    Pnm = Pj.Name
    For Each C In CmpAyzP(Pj, WhStr)
        Push MdDryzV, MdDr(CvMd(C))
    Next
Next
End Function
Function MdDr(A As CodeModule) As Variant()

End Function

Sub SavVbe(A As Vbe)
Dim P As VBProject
For Each P In A.VBProjects
    SavPj P
Next
End Sub

Function VisWinCntz%(A As Vbe)
VisWinCntz = NItrPrpTrue(A.Windows, "Visible")
End Function

Function DrOfMthLinyzV(A As Vbe, Optional WhStr$) As Variant()
Dim P
For Each P In PjItr(A, WhStr)
    PushObjAy DrOfMthLinyzV, DryOfMthLinzP(CvPj(P), WhStr)
Next
End Function
Function PjAy(A As Vbe, Optional WhStr$, Optional NmPfx$) As VBProject()
Dim P As VBProject, W As WhNm
Set W = WhNmzStr(WhStr, NmPfx)
For Each P In A.VBProjects
    If HitNm(P.Name, W) Then
        PushObj PjAy, P
    End If
Next
End Function

Function ItrwNmStr(Itr, WhStr$, Optional NmPfx$)
Stop '
End Function
Function ItrwNm(Itr, B As WhNm)
Stop '
End Function

Function PjItr(A As Vbe, Optional WhStr$, Optional NmPfx$)
If WhStr = "" Then
    Set PjItr = A.VBProjects
Else
    Asg PjAy(A, WhStr, NmPfx), PjItr
End If
End Function

Property Get PjfSyInVbe() As String()
PjfSyInVbe = PjfSyzV(CVbe)
End Property

Function PjfSyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNonBlank PjfSyzV, Pjf(P)
Next
End Function

Function PjNyInVbe(Optional WhStr$, Optional NmPfx$) As String()
PjNyInVbe = PjNyzV(CVbe, WhStr, NmPfx)
End Function

Function PjNyzV(A As Vbe, Optional WhStr$, Optional NmPfx$) As String()
Dim P
For Each P In PjItr(A, WhStr, NmPfx)
    PushI PjNyzV, CvPj(P).Name
Next
End Function

Function FstQPj(A As Vbe) As VBProject
Dim I
For Each I In A.VBProjects
    If FstChr(CvPj(I).Name) = "Q" Then
        Set FstQPj = I
        Exit Function
    End If
Next
End Function

Function MthWbzV(A As Vbe) As Workbook
Set MthWbzV = ShwWb(WbzWs(MthWszV(A)))
End Function

Function SrtRpt() As String()
SrtRpt = SrtRptzV(CVbe)
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

Private Sub ZZ_VbeFunPfx()
'D Vbe_MthPfx(CVbe)
End Sub

Private Sub ZZ_MthNyzV()
'Brw MthNyzV(CVbe)
End Sub

Private Sub ZZ_MthNyzVWh()
'Brw MthNyzV(CVbe)
End Sub

Private Sub ZZ()
Dim A
Dim B As Vbe
Dim C$
Dim D As Boolean
Dim E As WhPjMth
Dim F As WhNm
Dim XX
CvVbe A
End Sub

