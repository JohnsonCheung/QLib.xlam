Attribute VB_Name = "MIde_Exp_SrcPth"
Option Explicit
Function SrcpPj$()
SrcpPj = Srcp(CurPj)
End Function

Function SrcpzCmp$(A As VBComponent)
SrcpzCmp = Srcp(PjzCmp(A))
End Function

Function SrcpzPjf$(Pjf$)
SrcpzPjf = AddFdrApEns(Pth(Pjf), ".Src", Fn(Pjf))
End Function

Sub EnsSrcp(A As VBProject)
EnsPthzAllSeg Srcp(A)
End Sub

Function SrcpzDistPj$(DistPj As VBProject)
Dim P$: P = PjPth(DistPj)
SrcpzDistPj = AddFdrAp(PthUp(P, 2), ".Src", Fdr(P))
End Function

Function PthRmvFdr$(Pth$)
PthRmvFdr = BefRev(RmvPthSfx(Pth), PthSep) & PthSep
End Function

Function FfnUp$(Ffn$)
FfnUp = PthRmvFdr(Pth(Ffn$))
End Function
Function SrcpInPj$()
SrcpInPj = Srcp(CurPj)
End Function

Function Srcp$(A As VBProject)
Srcp = SrcpzPjf(Pjf(A))
End Function
Function SrcpOfPj$()
SrcpOfPj = SrcpzPj(CurPj)
End Function

Function IsSrcp(Pth$) As Boolean
IsSrcp = Fdr(ParPth(Pth)) = ".src"
End Function

Function SrcFn$(A As VBComponent)
SrcFn = A.Name & ".bas"
End Function

Sub ThwNotSrcp(Srcp$)
If Not IsSrcp(Srcp) Then Err.Raise 1, , "Not Srcp:" & vbCrLf & Srcp
End Sub

Function SrcFfn$(A As VBComponent)
SrcFfn = SrcpzCmp(A) & SrcFn(A)
End Function

Function IsInstScrp(Pth$) As Boolean
If Not IsPth(Pth) Then Exit Function
If Fdr(Pth) <> "Src" Then Exit Function
Dim P$: P = ParPth(Pth)
If Not IsDteTimStr(Fdr(P)) Then Exit Function
IsInstScrp = True
End Function

Sub ThwNotInstScrp(Pth$)
If Not IsInstScrp(Pth) Then Err.Raise 1, , "Not InstScrp(" & Pth & ")"
End Sub
