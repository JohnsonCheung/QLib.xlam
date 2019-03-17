Attribute VB_Name = "MIde_Md_Op_Rmv_Lines"
Option Explicit
Sub ClrMd(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("ClrMd: Md(?) of JnCrLf(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub RmvMdFTIxAy(A As CodeModule, B() As FTIx)
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

Sub RplMd(A As CodeModule, NewMdLines$)
RplMdLines A, MdLineszMd(A), NewMdLines, "Whole-Md"
End Sub

Sub RmvMdFTIx(A As CodeModule, FTIx As FTIx)
Dim FstLin$
FstLin = A.Lines(FTIx.FmNo, 1)
With FTIx
    If .Cnt = 0 Then Exit Sub
    A.DeleteLines .FmNo, .Cnt
    InfLin CSub, "Lines deleted", "Md Lno Cnt FstLin", MdNm(A), FTIx.FmNo, FTIx.Cnt, FstLin
End With
End Sub

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

