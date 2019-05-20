Attribute VB_Name = "QIde_Src_SLoc"
Option Explicit
Option Compare Text
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_StopLin."

Function SLocyzPP(P As VBProject, Patn$) As String()
Dim R As RegExp, O$(), Ly$(), Nm$, Md As CodeModule
Set R = RegExp(Patn)
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy SLocyzPP, SLocyzMR(C.CodeModule, R)
Next
End Function
Function SLocyzMR(A As CodeModule, R As RegExp) As String()
SLocyzMR = SLocyzSRN(Src(A), R, Mdn(A))
End Function

Function SLocyzSRN(Src$(), R, Mdn$) As String()
Dim L, Lno&, C1, C2
For Each L In Itr(Src)
'    PushI SLocyzSR, SLoc(Mdn, Lno, C1, C2)
Next
End Function
Function MdzSLoc(SLoc$) As CodeModule

End Function
Function RRCCzSLoc(SLoc$) As RRCC

End Function
Function JmpSLoc$(SLoc$)
'JmpMd MdzSLoc(SLoc)
JmpRRCC RRCCzSLoc(SLoc)
End Function

Function LnxszMdRe(A As CodeModule, R As RegExp) As Lnxs
Dim L$, J&
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If R.Test(L) Then
        PushLnx LnxszMdRe, Lnx(J - 1, ContLinzML(A, J))
    End If
Next
End Function


