Attribute VB_Name = "QIde_Loc"
Option Explicit
Private Const CMod$ = "MIde_Loc."
Private Const Asm$ = "QIde"
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
            PushMdPos MthPoses, MdPos(M, LinPos(Lno, SubStrPos(MthLin, Mthn)))
        Next
    End With
Next
End Function

Function LocLyzPatn(Patn$) As String()
LocLyzPatn = LocLyzPPatn(CPj, Patn)
End Function

Function LocLyzPPatn(P As VBProject, Patn$) As String()
LocLyzPPatn = SywPatn(SrczP(P), Patn)
End Function

Function LocLyzRlzP(P As VBProject, Re As RegExp) As String()
Dim C As VBComponent
For Each C In P.VBComponents
'    PushAy LocLyzPjRe, LocLyzMRe(C.CodeModule, Re)
Next
End Function

