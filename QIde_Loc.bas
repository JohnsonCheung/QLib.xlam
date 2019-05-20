Attribute VB_Name = "QIde_Loc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Loc."
Private Const Asm$ = "QIde"
Sub LisPatn(Patn$)
D LocLyzPatn(Patn)
End Sub
Function MthPoses(Mthn) As MdPoses
Dim R As Rel: 'Set R = MthPfxSyMd
Dim Mdn, M As CodeModule, MthLnx
For Each Mdn In R.ParChd(Mthn).Itms
    Set M = Md(Mdn)
    With MthLnxszM(M)
        Dim J&
        For J = 0 To .N - 1
            Dim Lno, MthLin
                With .Ay(J)
                    Lno = .Ix + 1
                    MthLin = .Lin
                End With
'            PushMdPos MthPoses, MdPos(M, LinPos(Lno, SubStrPos(MthLin, Mthn)))
        Next
    End With
Next
End Function
Function LocLyzMR(M As CodeModule, Re As RegExp) As String()
LocLyzMR = LocLyzML(Mdn(M), LnxszSR(Src(M), Re))
End Function

Function LnxszSR(Src$(), Re As RegExp) As Lnxs
Dim Ix&, L
For Each L In Itr(Src)
    If Re.Test(L) Then PushLnx LnxszSR, Lnx(L, Ix)
    Ix = Ix + 1
Next
End Function

Function LocLyzML(Mdn$, L As Lnxs) As String()
Dim J&
For J = 0 To L.N - 1
Next
End Function

Function LocLyzPatn(Patn$) As String()
LocLyzPatn = LocLyzPR(CPj, RegExp(Patn))
End Function

Function LocLyzPR(P As VBProject, Re As RegExp) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy LocLyzPR, LocLyzMR(C.CodeModule, Re)
Next
End Function

