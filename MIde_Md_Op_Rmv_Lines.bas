Attribute VB_Name = "MIde_Md_Op_Rmv_Lines"
Option Explicit
Sub ClrMd(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("ClrMd: Md(?) of JnCrLf(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub RmvMdLineszFTIxAy(A As CodeModule, B() As FTIx)
If Not FTIxAyIsInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmNo, .Cnt
    End With
Next
End Sub

Function CntSzzMd$(A As CodeModule)
CntSzzMd = CntSzzLines(SrcLines(A))
End Function
Sub RplMd(A As CodeModule, NewMdLines$)
Dim OldLines$: OldLines = SrcLines(A)
Dim NewCntSz$: NewCntSz = CntSzzLines(NewMdLines)
If OldLines = NewMdLines Then
    InfoLin CSub, "No replace: Same", "MdNm LinesCnt", MdNm(A), NewCntSz
    Exit Sub
End If
Dim OldCntSz$: OldCntSz = CntSzzLines(OldLines)
ClrMd A
A.AddFromString NewMdLines
InfoLin CSub, "Replaced", "MdNm Type OldLin NewLin", MdNm(A), OldLines, NewCntSz
End Sub

Sub RmvMdFTIx(A As CodeModule, FTIx As FTIx)
With FTIx
    If .Cnt = 0 Then Exit Sub
    A.DeleteLines .FmNo, .Cnt
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
A = MthFTIxAyMdMth(Md("Md_"), "XXX")
RmvMdFtLinesIxAy Md("Md_"), A
End Sub

