Attribute VB_Name = "QIde_Src_RetSrc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Src_Ret."
Private Const Asm$ = "QIde"

Function SrcLzMFi$(M As CodeModule, B As Fei)
If IsFeizEmp(B) Then Exit Function
'SrcLzMFi = A.Lines(B.FmNo, B.Cnt)
End Function

Function SrczMFi(M As CodeModule, B As Fei) As String()
'SrczMFi = SplitCrLf(SrcLzMFei(A, B))
End Function

Function SrczMRe(M As CodeModule, B As RegExp) As String()
Dim Ix&(): Ix = IxyzAyRe(Src(M), B)
Dim O$(), I, Md As CodeModule
Dim N$: N = Mdn(M)
If Si(Ix) = 0 Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdJmpLin ""?"",??' ?", N, I + 1, vbTab, M.Lines(I + 1, 1))
Next
SrczMRe = O
End Function

Function SrczPjPatn(P As VBProject, Patn$) As String()
SrczPjPatn = AwPatn(SrczP(P), Patn)
End Function



'
