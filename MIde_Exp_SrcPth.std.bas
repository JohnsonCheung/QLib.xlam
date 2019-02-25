Attribute VB_Name = "MIde_Exp_SrcPth"
Option Explicit
Function SrcPthPj$()
SrcPthPj = SrcPth(CurPj)
End Function

Function SrcPthzCmp$(A As VBComponent)
SrcPthzCmp = SrcPth(PjzCmp(A))
End Function

Function SrcPthzPjf$(Pjf)
SrcPthzPjf = AddFdr(Pth(Pjf), ".source", Fn(Pjf))
End Function

Function SrcPthEns$(A As VBProject)
SrcPthEns = PthEnsAll(SrcPth(A))
End Function

Function SrcPth$(A As VBProject)
SrcPth = SrcPthzPjf(Pjf(A))
End Function

Function IsSrcPth(Pth) As Boolean
IsSrcPth = Fdr(Pth) = "Src"
End Function

Function SrcFn$(A As VBComponent)
SrcFn = A.Name & SrcExt(A)
End Function

Sub ThwNotSrcPth(SrcPth)
If Not IsSrcPth(SrcPth) Then Err.Raise 1, , "Not SrcPth(" & SrcPth & ")"
End Sub

Function SrcExt$(A As VBComponent)
Select Case A.Type
Case vbext_ct_StdModule:   SrcExt = ".std.bas"
Case vbext_ct_ClassModule: SrcExt = ".cls.bas"
Case vbext_ct_Document:    SrcExt = ".doc.bas"
Case Else: Stop
End Select
End Function

Function SrcFfn$(A As VBComponent)
SrcFfn = SrcPthzCmp(A) & SrcFn(A)
End Function

Function IsSrcPthInst(Pth) As Boolean
If Not IsPth(Pth) Then Exit Function
If Fdr(Pth) <> "Src" Then Exit Function
Dim P$: P = ParPth(Pth)
If Not IsDteTimStr(Fdr(P)) Then Exit Function
IsSrcPthInst = True
End Function

Sub ThwNotSrcPthInst(Pth)
If Not IsSrcPthInst(Pth) Then Err.Raise 1, , "Not SrcPthInst(" & Pth & ")"
End Sub
