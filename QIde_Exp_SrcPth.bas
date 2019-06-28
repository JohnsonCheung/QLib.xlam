Attribute VB_Name = "QIde_Exp_SrcPth"
Option Explicit
Option Compare Text
Private Const CMod$ = "MIde_Exp_SrcPth."
Private Const Asm$ = "QIde"
Function SrcpPj$()
SrcpPj = Srcp(CPj)
End Function

Function SrcpzCmp$(A As VBComponent)
SrcpzCmp = Srcp(PjzC(A))
End Function

Function SrcpzPjf$(Pjf)
SrcpzPjf = AddFdrApEns(Pth(Pjf), ".Src", Fn(Pjf))
End Function

Sub EnsSrcp(P As VBProject)
EnsPthzAllSeg Srcp(P)
End Sub

Function SrcpzDistPj$(DistPj As VBProject)
Dim P$: P = Pjp(DistPj)
SrcpzDistPj = AddFdrAp(UpPth(P, 1), ".Src", Fdr(P))
End Function

Function SrcpP$()
SrcpP = Srcp(CPj)
End Function

Function Srcp$(P As VBProject)
Srcp = SrcpzPjf(Pjf(P))
End Function
Function HasExtss(Ffn, ExtSsLin) As Boolean
Dim E$: E = Ext(Ffn)
Dim Sy$(): Sy = SyzSS(ExtSsLin)
HasExtss = HasEleS(Sy, E)
End Function
Function IsSrcp(Pth) As Boolean
Dim F$: F = Fdr(Pth)
If Not HasExtss(F, ".xlam .accdb") Then Exit Function
IsSrcp = Fdr(ParPth(Pth)) = ".Src"
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

Function IsInstScrp(Pth) As Boolean
If Not IsPth(Pth) Then Exit Function
If Fdr(Pth) <> "Src" Then Exit Function
Dim P$: P = ParPth(Pth)
If Not IsTimStr(Fdr(P)) Then Exit Function
IsInstScrp = True
End Function

Sub ThwNotInstScrp(Pth)
If Not IsInstScrp(Pth) Then Err.Raise 1, , "Not InstScrp(" & Pth & ")"
End Sub
