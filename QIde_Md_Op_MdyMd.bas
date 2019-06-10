Attribute VB_Name = "QIde_Md_Op_MdyMd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Op_Rmv_Lines."
Private Const Asm$ = "QIde"

Sub ClrMd(M As CodeModule)
With M
    If .CountOfLines = 0 Then Exit Sub
    Debug.Print FmtQQ("ClrMd: Md(?) of Lines(?) is cleared", Mdn(M), .CountOfLines)
    .DeleteLines 1, .CountOfLines
    If .CountOfLines <> 0 Then Stop
End With
End Sub

Function CntSiStrzMd$(M As CodeModule)
CntSiStrzMd = CntSiStrzLines(SrcLines(M))
End Function

Sub DltLin(M As CodeModule, Lno, OldLin)
Dim LinFmMd$: LinFmMd = M.Lines(Lno, 1)
If LinFmMd <> OldLin Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(M), Lno, LinFmMd, OldLin
M.DeleteLines Lno, 1
End Sub

Sub DltLines(M As CodeModule, Lno, NLin, OldLines$)
Dim OldLinesFmMd$: OldLinesFmMd = M.Lines(Lno, NLin)
If OldLinesFmMd <> OldLines Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", Mdn(M), Lno, OldLinesFmMd, OldLines
M.DeleteLines Lno, NLin: Debug.Print FmtQQ("DltLines: Lno(?) NLin(?)")
End Sub

Sub DltLinzF(M As CodeModule, B As Feis)
If Not IsFeisInOrd(B) Then Thw CSub, "Given Feis is not in order", "Feis", LyzFeis(B)
Dim J%
For J = B.N - 1 To 0 Step -1
    With FCntzFei(B.Ay(J))
        M.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub

Sub DltLinzFei(M As CodeModule, B As Fei, OldLines$)
Stop
Dim FstLin
'FstLin = A.Lines(Fei.FmNo, 1)
With B
'    If .Cnt = 0 Then Exit Sub
'    A.DeleteLines .FmNo, .Cnt
End With
End Sub

Sub DltLinzFeis(M As CodeModule, B As Feis)
If Not IsFeisInOrd(B) Then Stop
Dim J&
For J = B.N - 1 To 0 Step -1
'    DltLinzFEITx B.Ay(J)
Next
End Sub

Sub InsLines(M As CodeModule, Lno, Lines$)
M.InsertLines Lno, Lines
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

Sub DltLinzD(M As CodeModule, L_OldL As Drs)
If JnSpc(L_OldL.Fny) <> "L OldL" Then Stop: Exit Sub
Stop
Dim B As Drs: B = SrtDrs(L_OldL, "-L")
Dim Dr
Stop
For Each Dr In Itr(B.Dry)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw CSub, "Md-Lin <> OldL", "Mdn Lno Md-Lin OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub

Sub InsLinzD(M As CodeModule, L_NewL As Drs)
If JnSpc(L_NewL.Fny) <> "L NewL" Then Stop: Exit Sub
Stop
Dim B As Drs: B = SrtDrs(L_NewL, "-L")
Dim Dr
Stop
For Each Dr In Itr(B.Dry)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw CSub, "Md-Lin <> OldL", "Mdn Lno Md-Lin OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub

Sub RplLin(M As CodeModule, L_NewL_OldL As Drs)
Dim B As Drs: B = L_NewL_OldL
If JnSpc(B.Fny) <> "L NewL OldL" Then Stop: Exit Sub
Dim Dr
For Each Dr In Itr(B.Dry)
    Dim L&: L = Dr(0)
    Dim OldL$: OldL = Dr(2)
    Dim NewL$: NewL = Dr(1)
    If M.Lines(L, 1) <> OldL Then Thw CSub, "Md-Lin <> OldL", "Mdn Lno Md-Lin OldL NewL", Mdn(M), L, M.Lines(L, 1), OldL, NewL
    M.ReplaceLine L, NewL
Next
End Sub

Sub RplLines(M As CodeModule, Lno, NLin, OldLines$, NewLines$)
DltLines M, Lno, NLin, OldLines
M.InsertLines Lno, NewLines
End Sub

Private Sub Z_DltLinzFeis()
Dim A As Feis
'A = MthFeiszMth(Md("Md_"), "XXX")
DltLinzFeis Md("Md_"), A
End Sub
