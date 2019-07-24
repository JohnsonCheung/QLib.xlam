Attribute VB_Name = "QDta_Fun_FmtDy"
Option Explicit
Option Compare Text

Private Sub Z_FmtDy()
Dim A$()
A = FmtDy(SampDy3, Fmt:=EiSSFmt, BrkCCIxy0:=Array(0))
A = FmtDy(SampDy3, Fmt:=EiSSFmt)
Dmp A
End Sub

Function WdtzCellCol%(CellCol$())
Dim O%, Cell
For Each Cell In Itr(CellCol)
    O = Max(O, WdtzCell(Cell))
Next
WdtzCellCol = O
End Function

Function WdtzCell%(Cell)
If IsLines(Cell) Then
    WdtzCell = WdtzLines(Cell)
Else
    WdtzCell = Len(Cell)
End If
End Function

Function WdtAyzCellDy(CellDy()) As Integer()
':CellDy: :Dy ! each dr is sy and  ele of the sy is a cell
Dim J&: For J = 0 To NColzDy(CellDy) - 1
    Push WdtAyzCellDy, WdtzCellCol(StrColzDy(CellDy, J))
Next
End Function

Function FmtDy(Dy(), _
Optional MaxColWdt% = 100, _
Optional BrkCCIxy0, _
Optional ShwZer As Boolean, _
Optional Fmt As EmTblFmt) _
As String()
If Si(Dy) = 0 Then Exit Function
Dim CellDy(): CellDy = XCellDy(Dy, ShwZer, MaxColWdt) ' :CellDy: ! each ele is a :Cell
Dim W%():                    W = WdtAyzCellDy(CellDy)
Dim WrpBdy$():          WrpBdy = XWrpBdy(CellDy, W, Fmt)
Dim Sep$:                  Sep = XSep(W, Fmt)
Dim BrkAy() As Boolean:  BrkAy = XBrkAy(Dy, BrkCCIxy0)
Dim BrkBdy$():          BrkBdy = XBrkBdy(WrpBdy, BrkAy, Sep)
Dim Bdy$():                Bdy = LyzLinesAy(BrkBdy)
                         FmtDy = Sy(Sep, Bdy, Sep)
End Function

Private Function XWrpBdy(CellDy(), W%(), Fmt As EmTblFmt) As String()
'Ret : sam ele as @Dy.  One ele of @Ret may be a lin or lines depending on any cell of Dr of @Dy has lines.
Dim S$: S = IIf(Fmt = EiSSFmt, " ", " | ")
Dim Q$: Q = IIf(Fmt = EiSSFmt, "", "| * |")
Dim Dr: For Each Dr In Itr(CellDy)
    PushI XWrpBdy, XLineszCellDr(CvSy(Dr), W, S, Q)
Next
'Insp "QDta_Fun_FmtDy.XWrpBdy", "Inspect", "Oup(XWrpBdy) CellDy W Fmt", XWrpBdy, "NoFmtr(())", W, "NoFmtr(EmTblFmt)": Stop
End Function

Private Function XNRowzCellDr%(CellDr$())
Dim NRow%, Cell: For Each Cell In CellDr
    NRow = Max(NRow, Si(SplitCrLf(Cell)))
Next
XNRowzCellDr = NRow
End Function

Private Function XColAyzCellDr(CellDr$()) As Variant()
Dim Cell: For Each Cell In CellDr
    PushI XColAyzCellDr, SplitCrLf(Cell)
Next
End Function

Private Function XSqzCellDr(CellDr$()) As Variant()
Dim NRow%: NRow = XNRowzCellDr(CellDr)
Dim NCol%: NCol = Si(CellDr)
Dim ColAy(): ColAy = XColAyzCellDr(CellDr)
Dim O(): ReDim O(1 To NRow, 1 To NCol)
Dim Col$(), ICol%, S$: For ICol = 0 To NCol - 1
    Col = ColAy(ICol)
    Dim IRow%: For IRow = 0 To UB(Col)
        S = Col(IRow)
        O(IRow + 1, ICol + 1) = S
    Next
Next
XSqzCellDr = O
End Function

Private Function XLineszCellDr$(CellDr$(), W%(), S$, Q$)
Dim Sq(): Sq = XSqzCellDr(CellDr)
Dim Sq1(): Sq1 = AlignSq(Sq, W)
Dim Dr, Lin$, O$(), IR%: For IR = 1 To UBound(Sq, 1)
    Dr = DrzSq(Sq1, IR)
    Lin = LinzDr(Dr, S, Q)
    PushI O, Lin
Next
XLineszCellDr = JnCrLf(O)
End Function

Private Function XSep$(W%(), Fmt As EmTblFmt)
Dim Sep$():  Sep = SepDr(W)
Dim S$:        S = IIf(Fmt = EiSSFmt, " ", "-|-")
Dim Q$:        Q = IIf(Fmt = EiSSFmt, "", "|-*-|")
            XSep = LinzDr(Sep, S, Q)
'Insp "QDta_Fun_FmtDy.XSep", "Inspect", "Oup(XSep) W Fmt", XSep, W, "NoFmtr(EmTblFmt)": Stop
End Function

Private Function XCellDy(Dy(), ShwZer As Boolean, MaxColWdt%) As Variant()
Dim J%
Dim Dr, StrDr$(): For Each Dr In Itr(Dy)
    StrDr = XCellDr(Dr, ShwZer, MaxColWdt)
    J = J + 1
    Push XCellDy, StrDr
Next
'Insp "QDta_Fun_FmtDy.XCellDy", "Inspect", "Oup(XCellDy) Dy ShwZer MaxColWdt", "NoFmtr(Variant())", "NoFmtr(())", ShwZer, "NoFmtr(% = 100)": Stop
End Function

Private Function XCellDr(Dr, ShwZer As Boolean, MaxWdt%) As String()
Dim V, S$: For Each V In Itr(Dr)
    S = Cell(V, ShwZer, MaxWdt)
    PushI XCellDr, S
Next
End Function

Private Function XBrkAy(Dy(), BrkCCIxy0) As Boolean()
'Ret : ! no or sam ele of @Dy.  Telling that row of @Dy should ins a brk aft the row. @@
Dim Ixy&(): Ixy = CvLngAy(BrkCCIxy0)
If Si(Ixy) = 0 Then Exit Function
Dim LasK: LasK = AwIxy(Dy(0), Ixy)
Dim CurK
Dim Dr: For Each Dr In Itr(Dy)
    CurK = AwIxy(Dr, Ixy)
           PushI XBrkAy, Not IsEqAy(CurK, LasK)
    LasK = CurK
Next
'Insp "QDta_Fun_FmtDy.XBrkAy", "Inspect", "Oup(XBrkAy) Dy BrkCCIxy0", "NoFmtr(Boolean())", "NoFmtr(())", BrkCCIxy0: Stop
End Function

Function XBrkBdy(Bdy$(), Brk() As Boolean, LinzSep$) As String()
If Si(Brk) = 0 Then XBrkBdy = Bdy: Exit Function
Dim L, J&: For Each L In Bdy
    If Brk(J) Then PushI XBrkBdy, LinzSep
    PushI XBrkBdy, L
    J = J + 1
Next
'Insp "QDta_Fun_FmtDy.XBrkBdy", "Inspect", "Oup(XBrkBdy) WrpBdy BrkAy Sep", XBrkBdy, WrpBdy, "NoFmtr(() As Boolean)", Sep: Stop
End Function

