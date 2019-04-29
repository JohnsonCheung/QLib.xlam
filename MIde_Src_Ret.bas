Attribute VB_Name = "MIde_Src_Ret"
Option Explicit

Function LineszMdFTIx$(A As CodeModule, B As FTIx)
If B.IsEmp Then Exit Function
LineszMdFTIx = A.Lines(B.FmNo, B.Cnt)
End Function

Function LyzMdFTIx(A As CodeModule, B As FTIx) As String()
LyzMdFTIx = SplitCrLf(LineszMdFTIx(A, B))
End Function

Function LyMdRe(A As CodeModule, B As RegExp) As String()
Dim Ix&(): Ix = IxAyzAyRe(Src(A), B)
Dim O$(), I, Md As CodeModule
Dim N$: N = MdNm(A)
If Si(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdJmpLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
LyMdRe = O
End Function

Function LyzPjPatn(A As VBProject, Patn$)
LyzPjPatn = SywPatn(SrczPj(A), Patn)
End Function


