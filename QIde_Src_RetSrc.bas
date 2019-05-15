Attribute VB_Name = "QIde_Src_RetSrc"
Option Explicit
Private Const CMod$ = "MIde_Src_Ret."
Private Const Asm$ = "QIde"

Function SrcLineszMFi$(A As CodeModule, B As Fei)
If IsFeizEmp(B) Then Exit Function
'SrcLineszMFi = A.Lines(B.FmNo, B.Cnt)
End Function

Function SrczMFi(A As CodeModule, B As Fei) As String()
'SrczMFi = SplitCrLf(SrcLineszMFei(A, B))
End Function

Function SrczMRe(A As CodeModule, B As RegExp) As String()
Dim Ix&(): Ix = IxyzAyRe(Src(A), B)
Dim O$(), I, Md As CodeModule
Dim N$: N = Mdn(A)
If Si(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdJmpLin ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
SrczMRe = O
End Function

Function SrczPjPatn(P As VBProject, Patn$) As String()
SrczPjPatn = SywPatn(SrczP(P), Patn)
End Function


