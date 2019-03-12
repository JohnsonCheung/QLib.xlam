Attribute VB_Name = "MIde_Src_Ret"
Option Explicit

Function LinesMd$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Function
LinesMd = A.Lines(1, A.CountOfLines)
End Function

Function LinesMdFTIx$(A As CodeModule, B As FTIx)
If B.IsEmp Then Exit Function
LinesMdFTIx = A.Lines(B.FmNo, B.Cnt)
End Function

Function LyMdFTIx(A As CodeModule, B As FTIx) As String()
LyMdFTIx = SplitCrLf(LinesMdFTIx(A, B))
End Function

Function LyMdRe(A As CodeModule, B As RegExp) As String()
Dim Ix&(): Ix = AyRe_IxAy(Src(A), B)
Dim O$(), I, Md As CodeModule
Dim N$: N = MdNm(A)
If Sz(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdJmpLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
LyMdRe = O
End Function

Function LyPjPatn(A As VBProject, Patn$)
LyPjPatn = AywPatn(SrczPj(A), Patn)
End Function

Function Src(A As CodeModule) As String()
Src = SplitCrLf(LinesMd(A))
End Function

Function SrcPj() As String()
SrcPj = SrczPj(CurPj)
End Function

Function SrcVbe() As String()
SrcVbe = SrczVbe(CurVbe)
End Function

Function SrczMdNm(MdNm$) As String()
SrczMdNm = Src(Md(MdNm))
End Function

Function SrczPj(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy SrczPj, Src(C.CodeModule)
Next
End Function

Function SrczVbe(A As Vbe) As String()
Dim P
For Each P In A.VBProjects
    PushIAy SrczVbe, SrczPj(CvPj(P))
Next
End Function

Property Get SrcMd() As String()
SrcMd = Src(CurMd)
End Property
