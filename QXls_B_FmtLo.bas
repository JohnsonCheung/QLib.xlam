Attribute VB_Name = "QXls_B_FmtLo"
Option Compare Text
Option Explicit
Private Const Asm$ = "QXls"
Private Const CMod$ = "MXls_Lo_Fmt."
Public Const DoczLof$ = "It is Ly with T1 LofT1nn"
Private Type A
    Lo As ListObject
    Fny() As String
    Lof() As String
End Type
Private A As A: 'Md lvl var set by !IniFmtLo!

Sub IniFmtLo(L As ListObject)
Set A.Lo = L
A.Fny = FnyzLo(L)
End Sub

Sub FmtLo(Lo As ListObject, Lof$())
Dim Fny$(): Fny = FnyzLo(Lo)
ThwIf_Er ErzLof(Lof, A.Fny), CSub
Dim D As Dictionary: Set D = DiT1qLyItr(Lof, "Ali Bdr Bet Cor Fml Fmt Lvl Tot Wdt Tit Nm Lbl")
Dim L
For Each L In D("Ali"): SetAli L: Next
For Each L In D("Bdr"): SetBdr L: Next
For Each L In D("Bet"): SetBet L: Next
For Each L In D("Cor"): SetCor L: Next
For Each L In D("Fml"): SetFml L: Next
For Each L In D("Fmt"): SetFmt L: Next
For Each L In D("Lvl"): SetLvl L: Next
For Each L In D("Tot"): SetTot L: Next
For Each L In D("Wdt"): SetWdt L: Next
SetTit Lo, LyzLyItr(D("Tit"))
SetLoNm Lo, T2(FstElezT1(A.Lof, "Nm"))
For Each L In D("Lbl"): SetLbl L: Next ' Must run Last
End Sub

'Ali -----------------------------------------------------------
Sub SetAli(LinOf_Ali_FldLikss)
'Fm LinOf_Ali_FldLikss : T1 is :Ali: Rst is FldLikss.  :Ali: is 'Left | Right | Center'
'Ret                   : align those col as stated in @FldLikss and in @LoFny as @Ali
Dim Ali$, Fny$(): XAsgT1Fny LinOf_Ali_FldLikss, Ali, Fny
Dim H As XlHAlign: H = XHAlign(Ali)
Dim F: For Each F In Itr(Fny)
    XCol(F).HorizontalAlignment = H
Next
End Sub

Private Sub XAsgT1Fny(LinOf_T1_FldLikss, OT1$, OFny$())
AsgT1Fny LinOf_T1_FldLikss, A.Fny, OT1, OFny
End Sub

Private Function XHAlign(Ali$) As XlHAlign
Select Case Ali
Case "Left": XHAlign = xlHAlignLeft
Case "Right": XHAlign = xlHAlignRight
Case "Center": XHAlign = xlHAlignCenter
Case Else: Inf CSub, "Invalid HAlignStr", "HAlignStr", Ali: Exit Function
End Select
End Function

'Bdr -----------------------------------------------------------
Sub SetBdr(LinOf_LRBoth_FldLikss)
'#  A.Fny                 :
'Fm LinOf_LRBoth_FldLikss : T1 is :Ali: Rst is FldLikss.  :LRBoth: is 'Left | Right | Both'
'Ret                   : align those col as stated in @FldLikss and in @LoFny as @Ali
Dim L$: L = LinOf_LRBoth_FldLikss
Dim LRBoth$, Fny$(): XAsgT1Fny L, LRBoth, Fny
Dim IsLeft As Boolean: IsLeft = HasEle(SyzSS("Left Both"), LRBoth)
Dim IsRight As Boolean: IsRight = HasEle(SyzSS("Right Both"), LRBoth)
Dim F: For Each F In Itr(Fny)
    If IsLeft Then BdrRgAy XColAy(Fny), xlEdgeLeft
    If IsRight Then BdrRgAy XColAy(Fny), xlEdgeRight
Next
End Sub

Private Function XColAy(Fny$()) As Range()
XColAy = ColAyzLoCny(A.Lo, Fny)
End Function
'Bet----------------------------------------------------------
Sub SetBet(LinOf_Sum_Fm_To)
Dim FSum$, FFm$, FTo$: AsgTTRst LinOf_Sum_Fm_To, FSum, FFm, FTo
XCol(FSum).Formula = FmtQQ("=Sum([?]:[?])", FFm, FTo)
End Sub
'Cor----------------------------------------------------------
Sub SetCor(LinOf_Cor_FldLikss)
Dim Cor$, Fny$(): XAsgT1Fny LinOf_Cor_FldLikss, Cor, Fny
Dim C&: C = Colr(Cor)
Dim F
For Each F In Itr(Fny)
    XCol(F).Color = C
Next
End Sub

'Fml----------------------------------------------------------
Private Sub SetFml(LinOf_Fld_Fml)
Dim F$, Fml$: AsgTRst LinOf_Fld_Fml, F, Fml
If Not HasEle(A.Fny, F) Then Inf CSub, "Fld not in Fny", "Fld-not-in-Fny Fml-Lin", F, A.Lo: Exit Sub
XCol(F).Formula = Fml
End Sub

'Fmt----------------------------------------------------------
Sub SetFmt(LinOf_Fmt_FldLikss)
Dim Fmt$, Fny$(): XAsgT1Fny LinOf_Fmt_FldLikss, Fmt, Fny
Dim F: For Each F In Itr(Fny)
    XCol(F).NumberFormat = Fmt
Next
End Sub

'Lbl----------------------------------------------------------
Sub SetLbl(LinOf_Fld_Lbl)
'Ret: Must run after forumla & Between
Dim Fld$, Lbl$:    AsgTRst LinOf_Fld_Lbl, Fld, Lbl
    If Not HasEle(A.Fny, Fld) Then Inf CSub, "Fld not in Fny", "Fld-with-er LblLin Fny", Fld, A.Lo, A.Fny: Exit Sub
Dim R1 As Range
Dim R2 As Range
Set R1 = XHdrCell(Fld)
Set R2 = CellAbove(R1)
SwapValzRg R1, R2
End Sub

'Lvl----------------------------------------------------------
Sub SetLvl(LinOf_Lvl_FldLikss)
Dim T1$, Fny$():       XAsgT1Fny LinOf_Lvl_FldLikss, T1, Fny
Dim Lvl As Byte: Lvl = T1
If Not IsBet(T1, "2", "8") And Len(T1) <> 1 Then Inf CSub, "Lvl should betwee 2 to 8", "Lvl LvlLin", Lvl, LinOf_Lvl_FldLikss: Exit Sub
Dim F: For Each F In Itr(Fny)
    XCol(F).OutlineLevel = Lvl
Next
End Sub

'Tot----------------------------------------------------------
Sub SetTot(LinOf_SACnt_FldLikss)
Dim SACnt$, Fny$():               XAsgT1Fny LinOf_SACnt_FldLikss, SACnt, Fny
Dim T As XlTotalsCalculation: T = XTotCalc(SACnt)
Dim F: For Each F In Itr(Fny)
    A.Lo.ListColumns(F).Total = T
Next
End Sub

Private Function XTotCalc(SACnt$) As XlTotalsCalculation
'Fm SACnt : "Sum | Avg | Cnt" @@
Dim O As XlTotalsCalculation
Select Case SACnt
Case "Sum": O = xlTotalsCalculationSum
Case "Avg": O = xlTotalsCalculationAverage
Case "Cnt": O = xlTotalsCalculationCount
Case Else: Inf CSub, "Invalid TotCalcStr", "TotCalcStr Valid-TotCalcStr", SACnt, "Sum Avg Cnt": Exit Function
End Select
XTotCalc = O
End Function

'Wdt----------------------------------------------------------
Private Sub SetWdt(LinOf_Wdt_FldLikss)
Dim Wdt$, Fny$(): XAsgT1Fny LinOf_Wdt_FldLikss, Wdt, Fny
Dim W%: W = Wdt
If Not IsBet(W, 5, 200) Then Inf CSub, "Invalid Wdt (should between 5 200)", "Wdt Lin", W, LinOf_Wdt_FldLikss: Exit Sub
Dim F: For Each F In Itr(Fny)
    RgzLc(A.Lo, F).EntireColumn.Width = W
Next
End Sub

'Tst-------------------------------------------------------------
Private Sub Z_FmtLo()
Dim Lo As ListObject, Fmtr() As String 'Lofr
'------------
Set Lo = SampLo
Fmtr = SampLof
GoSub Tst
Exit Sub
Tst:
    FmtLo Lo, Fmtr
    Return
End Sub

Private Sub Z_SetBdr()
Dim L$, Lo As ListObject, Fny$()
'--
Set Lo = SampLo
Fny = FnyzLo(Lo)
'--
GoSub T1
GoSub T2
Exit Sub
T1: L = "Left A B C": GoTo Tst
T2: L = "Left D E F": GoTo Tst
T3: L = "Right A B C": GoTo Tst
T4: L = "Center A B C": GoTo Tst
Tst:
    IniFmtLo Lo
    SetBdr L      '<=='
    Stop
    Return
End Sub

Private Sub Z()
QXls_B_FmtLo:
End Sub

'Fun===========================================================================
Function XCol(F) As Range
Set XCol = A.Lo.ListColumns(F).DataBodyRange.EntireColumn
End Function

Private Function XHdrCell(C) As Range
Set XHdrCell = RgA1(CellAbove(A.Lo.ListColumns(C).Range))
End Function

Sub FmtLoBStd(B As Workbook)
Dim S As Worksheet
For Each S In B.Sheets
    FmtLoSStd S
Next
End Sub
Sub FmtLoStd(L As ListObject)
FmtLo L, StdLof
End Sub
Sub FmtLoSStd(S As Worksheet)
Dim L As ListObject: For Each L In S.ListObjects
    FmtLo L, StdLof
Next
End Sub
Property Get StdLof() As String()

End Property

Sub AddFml(Lo As ListObject, ColNm$, Fml$)
Dim O As ListColumn
Set O = Lo.ListColumns.Add
O.Name = ColNm
O.DataBodyRange.Formula = Fml
End Sub

Sub AutoFit(A As ListObject, Optional MaxW = 100)
Dim C As Range: Set C = LoAllEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = EntRgC(C, J)
   If EntC.ColumnWidth > MaxW Then EntC.ColumnWidth = MaxW
Next
End Sub

Sub SetLoNmUgTbl(A As ListObject)
SetLoNm A, A.QueryTable.CommandText
End Sub

Sub SetLoNm(A As ListObject, Lon$)
If Lon <> "" Then
    If Not HasLo(WszLo(A), Lon) Then
        A.Name = Lon
    Else
        Inf CSub, "Lo"
    End If
End If
End Sub

Sub SetWrpLcc(A As ListObject, CC$, Optional Wrp As Boolean)
Dim C
For Each C In TermAy(CC)
    SetWrpLc A, C, Wrp
Next
End Sub

Sub SetWrpLc(A As ListObject, C, Optional Wrp As Boolean)
A.Lo.ListColumns(C).DataBodyRange.WrapText = Wrp
End Sub

Sub SetWdtLcc(A As ListObject, CC$, W)
Dim C
For Each C In TermAy(CC)
    SetWdtLc A, C, W
Next
End Sub

Sub SetWdtLc(L As ListObject, C, W)
EntColzLc(L, C).ColumnWidth = W
End Sub

Sub SetTotLnk(A As ListObject, C)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.Lo.ListColumns(C).DataBodyRange
Set Ws = WszRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Sub SetTit(A As ListObject, TitLy$())
Dim Sq(), R As Range
    Sq = XTitSq(TitLy, FnyzLo(A)): If Si(Sq) = 0 Then Exit Sub
    Set R = XTitAt(A, UBound(Sq(), 1))
Set R = RgzSq(Sq(), R)
XMgeTit R
BdrInside R
BdrAround R
End Sub

Private Sub XMgeTit(TitRg As Range)
Dim J%
For J = 1 To TitRg.Rows.Count
    XMgeTitH RgR(TitRg, J)
Next
For J = 1 To TitRg.Columns.Count
    XMgeTitV RgC(TitRg, J)
Next
End Sub

Private Sub XMgeTitH(TitRg As Range)
TitRg.Application.DisplayAlerts = False
Dim J%, C1%, C2%, V, LasV
LasV = RgRC(TitRg, 1, 1).Value
C1 = 1
For J = 2 To TitRg.Columns.Count
    V = RgRC(TitRg, 1, J).Value
    If V <> LasV Then
        C2 = J - 1
        If Not IsEmpty(LasV) Then
            RgRCC(TitRg, 1, C1, C2).MergeCells = True
        End If
        C1 = J
        LasV = V
    End If
Next
TitRg.Application.DisplayAlerts = True
End Sub

Private Sub XMgeTitV(A As Range)
Dim J%
For J = A.Rows.Count To 2 Step -1
    MgeCellAbove RgRC(A, J, 1)
Next
End Sub

Private Function XTitAt(Lo As ListObject, NTitRow%) As Range
Set XTitAt = RgRC(Lo.DataBodyRange, 0 - NTitRow, 1)
End Function

Private Function XTitSq(TitLy$(), LoFny$()) As Variant()
Dim Fny$()
Dim Col()
    Dim F$, I, Tit$
    For Each I In Fny
        F = I
        Tit = FstElezRmvT1(TitLy, F)
        If Tit = "" Then
            PushI Col, Sy(F)
        Else
            PushI Col, AyTrim(SplitVBar(Tit))
        End If
    Next
XTitSq = Transpose(SqzDy(Col))
End Function

Private Sub Z_XTitSq()
Dim TitLy$(), Fny$()
'----
Dim A$(), Act(), Ept()
'TitLy
    Erase A
    Push A, "A A1 | A2 11 "
    Push A, "B B1 | B2 | B3"
    Push A, "C C1"
    Push A, "E E1"
    TitLy = A
Fny = SyzSS("A B C D E")
Ept = XTitSq(TitLy, Fny)
    SetSqr Ept, 1, SyzSS("A1 B1 C1 D E1")
    SetSqr Ept, 2, Array("A2 11", "B2")
    SetSqr Ept, 3, Array(Empty, "B3")
GoSub Tst
Exit Sub
'---
'TitLy
    Erase A
    Push A, "ksdf | skdfj  |skldf jf"
    Push A, "skldf|sdkfl|lskdf|slkdfj"
    Push A, "askdfj|sldkf"
    Push A, "fskldf"
    TitLy = A
BrwSq XTitSq(TitLy, Fny)

Exit Sub
Tst:
    Act = XTitSq(TitLy, Fny)
    Ass IsEqSq(Act, Ept)
    Return
End Sub

Sub BrwSampLof()
Brw FmtLof(SampLof)
End Sub

Function FmtLof(Lof$()) As String()
FmtLof = FmtSpec(Lof, LofT1nn, 2)
End Function

Function FmtSpec(Spec$(), Optional T1nn$, Optional FmtFstNTerm% = 1) As String()
Dim mT1Ay$()
    If IsMissing(T1nn) Then
        mT1Ay = T1Ay(Spec)
    Else
        mT1Ay = TermAy(T1nn)
    End If
Dim O$()
    Dim T$, I
    For Each I In mT1Ay
        T = I
        PushIAy O, AwT1(Spec, T)
    Next
    Dim M$(): M = SyeT1Sy(Spec, mT1Ay)
    If Si(M) > 0 Then
        PushI O, FmtQQ("# Error: in not T1Ay(?)", TLin(mT1Ay))
        PushIAy O, M
    End If
FmtSpec = FmtSyzNTerm(O, FmtFstNTerm)
End Function

