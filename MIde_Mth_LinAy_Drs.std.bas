Attribute VB_Name = "MIde_Mth_LinAy_Drs"
Option Explicit
Const CMod$ = "MIde_MthLinAy."

Function MthLinAy(Optional WhStr$) As String()
MthLinAy = MthLinAyzVbe(CurVbe, WhStr)
End Function

Function MthLinAyzVbe(V As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In V.VBProjects
    PushIAy MthLinAyzVbe, MthLinAyzPj(P, WhStr)
Next
End Function
Function MthLinAyzPj1(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy MthLinAyzPj1, MthLinAyzMd(C.CodeModule)
Next
End Function

Function MthLinAyzMd(A As CodeModule, Optional WhStr$) As String()
MthLinAyzMd = MthLinAyzSrc(Src(A), WhStr)
End Function


Function MthLinAyzSrc(Src$(), Optional WhStr$) As String()
Dim O$(), J&, B As WhMth
Set B = WhMthzStr(WhStr)
For J = 0 To UB(Src)
    If HitMthLin(Src(J), B) Then
        PushI MthLinAyzSrc, ContLin(Src, J)
    End If
Next
End Function

Function MthNmFny() As String()
MthNmFny = SySsl("MthNm Ty Mdy Md MdTy Pj")
End Function

Function MthQLinAy() As String()
MthQLinAy = AyQSrt(MthQLinAyzbe(CurVbe))
End Function
Function MthQLinAyzbe(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MthQLinAyzbe, MthQLinAyPj(P)
Next
End Function
Function MthQLinAyMd(A As CodeModule) As String()
Dim P$
P = A.Parent.Collection.Parent.Name & "." & ShtCmpTy(A.Parent.Type) & "." & A.Parent.Name & "."
MthQLinAyMd = AyAddPfx(MthLinAyzSrc(Src(A)), P)
End Function
Function MthQLinAyPj(A As VBProject) As String()
Dim C As VBComponent
If A.Protection = vbext_pp_locked Then Exit Function
For Each C In A.VBComponents
    PushIAy MthQLinAyPj, MthQLinAyMd(C.CodeModule)
Next
End Function

Function MthLinAyzPj(A As VBProject, Optional WhStr$) As String()
Dim I
For Each I In ModItrPj(A, WhStr)
    PushIAy MthLinAyzPj, MthLinAyzMd(CvMd(I), WhStr)
Next
End Function

