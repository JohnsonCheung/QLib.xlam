Attribute VB_Name = "MIde_Loc"
Option Explicit
Function SubStrPos(A, SubStr) As Pos
Dim P&: P = InStr(A, SubStr)
SubStrPos = Pos(P, P + Len(SubStr) - 1)
End Function
Function MthPos(MthLin) As Pos
If IsMthLin(MthLin) Then
    MthPos = SubStrPos(MthLin, MthNm(MthLin))
End If
End Function

Function LocLyPatn(Patn$) As String()
LocLyPatn = LocLyPjPatn(CurPj, Patn)
End Function

Function LocLyPjPatn(A As VBProject, Patn$) As String()
LocLyPjPatn = AywPatn(SrczPj(A), Patn)
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

