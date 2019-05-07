Attribute VB_Name = "QIde_Cmd_Lis_Src"
Option Explicit
Private Const CMod$ = "MIde_Cmd_Lis_Src."
Private Const Asm$ = "QIde"
Sub LisSrc(Patn$)
LisSrczPj CurPj, Patn
End Sub
Function LyzMdRe(A As CodeModule, Re As RegExp) As String()
LyzMdRe = LyzLnxs(LnxszCmpRe(C))
End Function
Sub LisSrczPj(A As VBProject, Patn$)
Dim R As RegExp, O$(), Ly$(), Nm$, Md As CodeModule
Set R = RegExp(Patn)
Dim C As VBComponent
For Each C In A.VBComponents
    Nm = C.Name & "."
    Set Md = C.CodeModule
    Ly = LyzLnxs(LnxszMdRe(Md, R))
    Ly = AddPfxzSy(Ly, Nm)
    PushIAy O, Ly
Next
'Vc O
Stop
Vc AlignzBySepss(O, ".")
End Sub

Function LinzGoMdDnmLno$(MdDNm$, Lno&)
LinzGoMdDnmLno = FmtQQ("MdNmLnoGo ""?"",?", MdDNm, Lno)
End Function
Function LyzLnxs(A As Lnxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzLnxs, LinzLnx(A.Ay(J))
Next
End Function

Function LinzLnx$(A As Lnx)
With A
LinzLnx = .Ix & "." & .Lin
End With
End Function
Function LnxszMdRe(A As CodeModule, R As RegExp) As Lnxs
Dim L$, J&
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If R.Test(L) Then
        PushLnx LnxszMdRe, Lnx(J - 1, ContLinzMd(A, J))
    End If
Next
End Function

