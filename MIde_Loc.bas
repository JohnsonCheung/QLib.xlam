Attribute VB_Name = "MIde_Loc"
Option Explicit
Function MthPos(MthNm) As MdPos()
Dim R As Rel: Set R = RelOf_MthNm_To_MdNy_InPj
Dim MdNm, M As CodeModule, MthLnx
For Each MdNm In R.ParChd(MthNm).Itms
    Set M = Md(MdNm)
    For Each MthLnx In MthLnxAyzMd(M)
        Dim Lno&, MthLin$
            With CvLnx(MthLnx)
            Lno = .Ix + 1
            MthLin = .Lin
            End With
        Dim P As Pos
            Pos
            Set P = SubStrPos(MthLin, MthNm)
        PushObj MthPos, MdPos(M, LinPos(Lno, SubStrPos(MthLin, MthNm)))
    Next
Next
End Function

Function LocLyPatn(Patn$) As String()
LocLyPatn = LocLyzPjPatn(CurPj, Patn)
End Function

Function LocLyzPjPatn(A As VBProject, Patn$) As String()
LocLyzPjPatn = AywPatn(SrczPj(A), Patn)
End Function

Function CurLocLyPjRe(Re_Or_Patn) As String()

End Function

Function LocLyPjRe(A As VBProject, Re As RegExp) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy LocLyPjRe, LocLyMdRe(C.CodeModule, Re)
Next
End Function

Function LocLyMdRe(A As CodeModule, Re As RegExp) As String()

End Function

