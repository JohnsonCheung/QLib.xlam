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
Function LinzFei$(A As Fei)
With A
LinzFei = "FmEndIx " & .FmIx & " " & .EIx
End With
End Function
Function LyzFeis(A As Feis) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzFeis, J & " " & LinzFei(A.Ay(J))
Next
End Function

Sub DltLinzF(A As CodeModule, B As Feis)
If Not IsFeisInOrd(B) Then Thw CSub, "Given Feis is not in order", "Feis", LyzFeis(B)
Dim J%
For J = B.N - 1 To 0 Step -1
    With FCntzFei(B.Ay(J))
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Function CntSiStrzMd$(A As CodeModule)
CntSiStrzMd = CntSiStrzLines(SrcLines(A))
End Function

Sub RplMd(A As RplgMd)
ClrMd A.Md
A.Md.InsertLines 1, A.NewLines
End Sub

Sub DltLinzFei(A As CodeModule, B As Fei, OldLines$)
Stop
Dim FstLin
'FstLin = A.Lines(Fei.FmNo, 1)
With B
'    If .Cnt = 0 Then Exit Sub
'    A.DeleteLines .FmNo, .Cnt
End With
End Sub

Sub DltLinzFeis(A As CodeModule, B As Feis)
If Not IsFeisInOrd(B) Then Stop
Dim J&
For J = B.N - 1 To 0 Step -1
'    DltLinzFEITx B.Ay(J)
Next
End Sub

Private Sub Z_DltLinzFeis()
Dim A As Feis
'A = MthFeiszMth(Md("Md_"), "XXX")
DltLinzFeis Md("Md_"), A
End Sub

Sub MdyMdzMM(A As CodeModule, B As Mdyg)
With B
Select Case .Act
Case EiIns: InsLinzM A, .Ins
Case EiDlt: DltLinzM A, .Dlt
Case Else: Thw CSub, "Unexpected Act.  Should be Ins or Rpl only", "Act", Act
End Select
End With
End Sub
Sub InsLinzM(A As CodeModule, B As Insg)
InsLines A, B.Lno, B.Lin
End Sub
Sub DltLinzM(A As CodeModule, B As Dltg)
If A.Lines(B.Lno, 1) <> B.Lin Then Thw CSub, "Ept-Lin <> Act-Lin", "Md At-Lno# Ept-Lin Act-Lin", Mdn(A), B.Lno, B.Lin, A.Lines(B.Lno, 1)
A.DeleteLines B.Lno
End Sub

Sub InsLines(A As CodeModule, Lno, Lines$)
A.InsertLines Lno, Lines
End Sub

Sub RplLines(A As CodeModule, Lno, NLin, OldLines$, NewLines$)
DltLines A, Lno, NLin, OldLines
InsLines A, Lno, NewLines
End Sub

Sub DltLines(A As CodeModule, Lno, NLin, OldLines$)
Dim OldLinesFmMd$: OldLinesFmMd = A.Lines(Lno, NLin)
If OldLinesFmMd <> OldLines Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(A), Lno, OldLinesFmMd, OldLines
A.DeleteLines Lno, NLin
End Sub

Sub DltLin(A As CodeModule, Lno, OldLin)
Dim LinFmMd$: LinFmMd = A.Lines(Lno, 1)
If LinFmMd <> OldLin Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(A), Lno, LinFmMd, OldLin
A.DeleteLines Lno, 1
End Sub


