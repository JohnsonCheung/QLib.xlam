Attribute VB_Name = "MIde_Md_Op_Rmv_Lines"
Option Explicit
Sub ClrMd(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("ClrMd: Md(?) of JnCrLf(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub MdRmvFTIxAy(A As CodeModule, B() As FTIx)
If Not FTIxAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmNo, .Cnt
    End With
Next
End Sub

Function CntSzStrzMd$(A As CodeModule)
CntSzStrzMd = CntSzStrzLines(SrcLines(A))
End Function

Function MdLineszMd(A As CodeModule) As MdLines
Set MdLineszMd = MdLines(1, SrcLines(A))
End Function

Function MdRpl(A As CodeModule, NewMdLines$) As CodeModule
Set MdRpl = MdRplLines(A, MdLineszMd(A), NewMdLines, "Whole-Md")
End Function

Function MdRmvFTIx(A As CodeModule, FTIx As FTIx) As CodeModule
Dim FstLin$
FstLin = A.Lines(FTIx.FmNo, 1)
With FTIx
    If .Cnt = 0 Then Exit Function
    A.DeleteLines .FmNo, .Cnt
    InfLin CSub, "Lines deleted", "Md Lno Cnt FstLin", MdNm(A), FTIx.FmNo, FTIx.Cnt, FstLin
End With
End Function

Sub RmvMdFtLinesIxAy(A As CodeModule, B() As FTIx)
If Not FTIxAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmNo, .Cnt
    End With
Next
End Sub


Private Sub Z_RmvMdFtLinesIxAy()
Dim A() As FTIx
A = MthFTIxAyzMth(Md("Md_"), "XXX")
RmvMdFtLinesIxAy Md("Md_"), A
End Sub

