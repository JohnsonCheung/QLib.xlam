Attribute VB_Name = "MIde_Mth_Nm_DNm"
Option Explicit

Private Sub Z_MthDNySrc()
BrwAy MthDNySrc(SrcMd)
End Sub

Function MthDNy(Optional WhStr$) As String()
MthDNy = MthDNyVbe(CurVbe, WhStr)
End Function

Function MthDNyVbe(A As Vbe, Optional WhStr$) As String()
Dim P As VBProject
For Each P In PjItr(A, WhStr)
    PushIAy MthDNyVbe, MthDNyPj(P, WhStr)
Next
End Function

Function MthDNyMd(A As CodeModule, Optional WhStr$) As String()
MthDNyMd = MthDNySrc(Src(A), WhStr)
End Function

Function MthDNySrc(Src$(), Optional WhStr$) As String()
Dim L
For Each L In Itr(Src)
    PushNonBlankStr MthDNySrc, MthDNmLin(L)
Next
End Function

