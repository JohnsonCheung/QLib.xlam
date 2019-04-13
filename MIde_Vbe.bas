Attribute VB_Name = "MIde_Vbe"
Option Explicit

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function
Sub DmpPjIsSav()
DmpDrs PjIsSavDrszVbe(CurVbe)
End Sub
Function PjIsSavDryzVbe(A As Vbe) As Variant()
Dim I As VBProject
For Each I In A.VBProjects
    PushI PjIsSavDryzVbe, Array(I.Saved, I.Name, I.BldFileName)
Next
End Function
Function PjIsSavDrszVbe(A As Vbe) As Drs
Set PjIsSavDrszVbe = Drs("IsSav PjNm BldFfn", PjIsSavDryzVbe(A))
End Function

Function Vbe_Pj(A As Vbe, PjNm$) As VBProject
Set Vbe_Pj = A.VBProjects(PjNm)
End Function

Function PjzPjfVbe(Vbe As Vbe, PjFil) As VBProject
Dim I As VBProject
For Each I In Vbe.VBProjects
    If Pjf(I) = PjFil Then Set PjzPjfVbe = I: Exit Function
Next
End Function

Function MdDryzVbe(A As Vbe, Optional WhStr$) As Variant()
Dim P, C, Pnm$, Pj As VBProject
For Each P In PjItr(A, WhStr)
    Set Pj = P
    Pnm = Pj.Name
    For Each C In CmpAyzPj(Pj, WhStr)
        Push MdDryzVbe, MdDr(CvMd(C))
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

Sub CompileVbe(A As Vbe)
DoItrFun A.VBProjects, "PjCompile"
End Sub
Function MthLinDryzVbe(A As Vbe, Optional WhStr$) As Variant()
Dim P
For Each P In PjItr(A, WhStr)
    PushObjAy MthLinDryzVbe, MthLinDryzPj(CvPj(P), WhStr)
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

Property Get PjfAyInVbe() As String()
PjfAyInVbe = PjfAyzVbe(CurVbe)
End Property

Function PjfAyzVbe(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushNonBlankStr PjfAyzVbe, Pjf(P)
Next
End Function

Function PjNyInVbe(Optional WhStr$, Optional NmPfx$) As String()
PjNyInVbe = PjNyzVbe(CurVbe, WhStr, NmPfx)
End Function

Function PjNyzVbe(A As Vbe, Optional WhStr$, Optional NmPfx$) As String()
Dim P
For Each P In PjItr(A, WhStr, NmPfx)
    PushI PjNyzVbe, CvPj(P).Name
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

Function MthWbVbe(A As Vbe) As Workbook
Set MthWbVbe = WbVis(WbzWs(MthWszVbe(A)))
End Function

Function VbeSrtRpt() As String()
VbeSrtRpt = SrtRptVbe(CurVbe)
End Function

Function HasVbeBar(A As Vbe, Nm$) As Boolean
HasVbeBar = HasItn(A.CommandBars, Nm)
End Function

Function Vbe_HasPj(A As Vbe, PjNm) As Boolean
Vbe_HasPj = HasItn(A.VBProjects, PjNm)
End Function

Function HasPjfVbe(A As Vbe, Ffn) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If Pjf(P) = Ffn Then HasPjfVbe = True: Exit Function
Next
End Function

Function SrtRptVbe(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy SrtRptVbe, SrtRptzPj(P)
Next
End Function

Private Sub ZZ_VbeFunPfx()
'D Vbe_MthPfx(CurVbe)
End Sub

Private Sub ZZ_MthNyzVbe()
'Brw MthNyzVbe(CurVbe)
End Sub

Private Sub ZZ_MthNyzVbeWh()
'Brw MthNyzVbe(CurVbe)
End Sub

Private Sub ZZ()
Dim A
Dim B As Vbe
Dim C$
Dim D As Boolean
Dim E As WhPjMth
Dim F As WhNm
Dim xx
CvVbe A
PjzPjfVbe B, A
'VbePjf B, A
PjzPjfVbe B, A
MthWbVbe B
HasVbeBar B, C
Vbe_HasPj B, A
HasPjfVbe B, A
End Sub

Private Sub Z()
End Sub
