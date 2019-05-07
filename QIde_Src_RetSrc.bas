Attribute VB_Name = "QIde_Src_RetSrc"
Option Explicit
Private Const CMod$ = "MIde_Src_Ret."
Private Const Asm$ = "QIde"

Function SrcLineszMdFTIx$(A As CodeModule, B As FTIx)
If B.IsEmp Then Exit Function
SrcLineszMdFTIx = A.Lines(B.FmNo, B.Cnt)
End Function

Function SrczMdFTIx(A As CodeModule, B As FTIx) As String()
LyzMdFTIx = SplitCrLf(LineszMdFTIx(A, B))
End Function

Function SrczMdRe(A As CodeModule, B As RegExp) As String()
Dim Ix&(): Ix = IxAyzAyRe(Src(A), B)
Dim O$(), I, Md As CodeModule
Dim N$: N = MdNm(A)
If Si(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdJmpLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
LyMdRe = O
End Function

Function SrczPjPatn(A As VBProject, Patn$) As String()
SrczPjPatn = SywPatn(SrczPj(A), Patn)
End Function


