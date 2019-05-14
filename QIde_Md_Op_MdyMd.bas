Attribute VB_Name = "QIde_Md_Op_MdyMd"
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Rmv_Lines."
Private Const Asm$ = "QIde"
Sub ClrMd(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("ClrMd: Md(?) of JnCrLf(?) is cleared", Mdn(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub
Function LinzFEIx$(A As FEIx)
With A
LinzFEIx = "FmEndIx " & .FmIx & " " & .EIx
End With
End Function
Function LyzFEIxs(A As FEIxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzFEIxs, J & " " & LinzFEIx(A.Ay(J))
Next
End Function

Sub DltLinzF(A As CodeModule, B As FEIxs)
If Not IsFEIxsInOrd(B) Then Thw CSub, "Given FEIxs is not in order", "FEIxs", LyzFEIxs(B)
Dim J%
For J = B.N - 1 To 0 Step -1
    With FCntzFEIx(B.Ay(J))
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Function CntSiStrzMd$(A As CodeModule)
CntSiStrzMd = CntSiStrzLines(SrcLines(A))
End Function

Sub RplMd(A As CodeModule, NewMdLines$)
'RplLines A, MdLineszMd(A), NewMdLines
End Sub

Sub DltLinzFEIx(A As CodeModule, B As FEIx, OldLines$)
Stop
Dim FstLin
'FstLin = A.Lines(FEIx.FmNo, 1)
With B
'    If .Cnt = 0 Then Exit Sub
'    A.DeleteLines .FmNo, .Cnt
End With
End Sub

Sub DltLinzFEIxs(A As CodeModule, B As FEIxs)
If Not IsFEIxsInOrd(B) Then Stop
Dim J&
For J = B.N - 1 To 0 Step -1
'    DltLinzFEITx B.Ay(J)
Next
End Sub

Private Sub Z_DltLinzFEIxs()
Dim A As FEIxs
'A = MthFEIxszMth(Md("Md_"), "XXX")
DltLinzFEIxs Md("Md_"), A
End Sub

Sub MdyMdzMM(A As CodeModule, B As Mdyg)
With B
Select Case .Act
Case EiIns: InsLinzMI A, .Ins
Case EiDlt: DltLinzMD A, .Dlt
Case EiRpl: RplLinzMR A, .Rpl
Case Else: Thw CSub, "Unexpected Act.  Should be Ins or Rpl only", "Act", Act
End Select
End With
End Sub
Sub InsLinzMI(A As CodeModule, Ins As InsgLin)

End Sub
Sub DltLinzMD(A As CodeModule, Dlt As DltgLin)

End Sub
Sub RplLinzMR(A As CodeModule, Rpl As RplgLin)

End Sub
Sub InsLinzInsg(A As CodeModule, B As InsgLin)
InsLin A, B.Lno, B.Lines
End Sub

Sub InsLines(A As CodeModule, Lno, Lines$)
A.InsertLines Lno, Lin
End Sub

Sub RplLin(A As CodeModule, Lno, OldLin$, NewLin$)
If A.Lines(Lno, 1) <> OldLines Then Thw CSub, "Md-Lin <> OldLno", "Md Lno Md-Lin OldLin NewLin", Mdn(A), Lno, A.Lines(Lno, 1), OldLines
A.ReplaceLine Lno, NewLines
End Sub

Sub RplLines(A As CodeModule, Lno, NLin, OldLines$, NewLines$)
DltLines A, Lno, NLin, OldLines
InsLines A, Lno, NewLines
End Sub

Sub DltLines(A As CodeModule, Lno, NLin, OldLines$)
Dim LinesFmMd$: LinesFmMd = A.Lines(Lno, NLin)
If LinesFmMd <> OldLines Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(A), Lno, LinFmMd, OldLin
A.DeleteLines Lno, NLin
End Sub

Sub DltLin(A As CodeModule, Lno, OldLin)
Dim LinFmMd$: LinFmMd = A.Lines(Lno, 1)
If LinFmMd <> OldLin Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(A), Lno, LinFmMd, OldLin
A.DeleteLines Lno, 1
End Sub


