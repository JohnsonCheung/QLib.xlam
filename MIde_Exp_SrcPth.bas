Attribute VB_Name = "MIde_Exp_SrcPth"
Option Explicit
Function SrcpPj$()
SrcpPj = Srcp(CurPj)
End Function

Function SrcpzCmp$(A As VBComponent)
SrcpzCmp = Srcp(PjzCmp(A))
End Function

Function SrcpzPjf$(Pjf)
SrcpzPjf = AddFdr(Pth(Pjf), ".src", Fn(Pjf))
End Function

Function SrcpzEns$(A As VBProject)
SrcpzEns = PthEnsAll(Srcp(A))
End Function

Function SrcpzDistPj$(DistPj As VBProject)
Dim P$: P = PjPth(DistPj)
SrcpzDistPj = AddFdr(PthUp(P, 2), ".src", Fdr(P))
End Function

Function PthRmvFdr$(Pth)
PthRmvFdr = StrBefRev(PthRmvSfx(Pth), PthSep) & PthSep
End Function

Function FfnUp$(Ffn)
FfnUp = PthRmvFdr(Pth(Ffn))
End Function

Function Srcp$(A As VBProject)
Srcp = SrcpzPjf(Pjf(A))
End Function

Function IsSrcp(Pth) As Boolean
IsSrcp = Fdr(ParPth(Pth)) = ".src"
End Function

Function SrcFn$(A As VBComponent)
SrcFn = A.Name & ".bas"
End Function

Sub ThwNotSrcp(Srcp)
If Not IsSrcp(Srcp) Then Err.Raise 1, , "Not Srcp:" & vbCrLf & Srcp
End Sub

Function SrcFfn$(A As VBComponent)
SrcFfn = SrcpzCmp(A) & SrcFn(A)
End Function

Function IsSrcpInst(Pth) As Boolean
If Not IsPth(Pth) Then Exit Function
If Fdr(Pth) <> "Src" Then Exit Function
Dim P$: P = ParPth(Pth)
If Not IsDteTimStr(Fdr(P)) Then Exit Function
IsSrcpInst = True
End Function

Sub ThwNotSrcpInst(Pth)
If Not IsSrcpInst(Pth) Then Err.Raise 1, , "Not SrcpInst(" & Pth & ")"
End Sub
