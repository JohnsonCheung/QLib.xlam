Attribute VB_Name = "QXls_Lo_FmtLo"
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
Private A As A

Sub FmtLo(Lo As ListObject, Lof$())
Set A.Lo = Lo
A.Fny = FnyzLo(Lo)
A.Lof = Lof
ThwIf_Er ErzLof(Lof, A.Fny), CSub
Dim L
For Each L In WItr("Ali"): WFmtAli CStr(L): Next
For Each L In WItr("Bdr"): WFmtBdr CStr(L): Next
For Each L In WItr("Bet"): WFmtBet CStr(L): Next
For Each L In WItr("Cor"): WFmtCor CStr(L): Next
For Each L In WItr("Fml"): WFmtFml CStr(L): Next
For Each L In WItr("Fmt"): WFmtFmt CStr(L): Next
For Each L In WItr("Lvl"): WFmtLvl CStr(L): Next
For Each L In WItr("Tot"): WFmtTot CStr(L): Next
For Each L In WItr("Wdt"): WFmtWdt CStr(L): Next
SetLoTit Lo, WLy("Tit")
SetLoNm Lo, T2(FstElezT1(A.Lof, "Nm"))
For Each L In WItr("Lbl"): WFmtLbl CStr(L): Next ' Must run Last
End Sub

'Ali -----------------------------------------------------------
Private Sub WFmtAli(L$)
Dim T1$, Fny$(): WAsgT1zFny T1, Fny, L
Dim H As XlHAlign: H = HAlign(T1)
Dim F$, I
For Each I In Itr(Fny)
    F = I
    WRg(F).HorizontalAlignment = H
Next
End Sub

Private Function HAlign(S) As XlHAlign
Select Case S
Case "Left": HAlign = xlHAlignLeft
Case "Right": HAlign = xlHAlignRight
Case "Center": HAlign = xlHAlignCenter
Case Else: Inf CSub, "Invalid HAlignStr", "HAlignStr", S: Exit Function
End Select
End Function

Private Sub WAsgT1zFny(OT1, OFny$(), Lin)
Dim FldLikAy$(): AsgT1FldLikAy OT1, FldLikAy, Lin
OFny = WFnyzFldLikAy(FldLikAy)
End Sub
Private Function WFnyzFldLikAy(FldLikAy$()) As String()
Dim F$, I
For Each I In A.Fny
    F = I
    If HitLikAy(F, FldLikAy) Then PushI WFnyzFldLikAy, F
Next
End Function
'Bdr -----------------------------------------------------------
Private Sub WFmtBdr(L$)
Dim LRBoth$, Fny$(): WAsgT1zFny LRBoth, Fny, L
Dim IsLeft As Boolean: IsLeft = HasEle(SyzSS("Left Both"), LRBoth)
Dim IsRight As Boolean: IsRight = HasEle(SyzSS("Right Both"), LRBoth)
If IsLeft Then BdrRgAy WRgAy(Fny), xlEdgeLeft
If IsRight Then BdrRgAy WRgAy(Fny), xlEdgeRight
End Sub
'Bet----------------------------------------------------------
Private Sub WFmtBet(L$)
Dim Fld$, FmFld$, ToFld$
    AsgN2tRst L, Fld, FmFld, ToFld
    Dim IsEr As Boolean
    If Not HasEle(A.Fny, Fld) Then IsEr = True
    If Not HasEle(A.Fny, FmFld) Then IsEr = True
    If Not HasEle(A.Fny, ToFld) Then IsEr = True
    If IsEr Then Inf CSub, "Bet-Lin error, the Fld,FmFld,ToFld not all in Fny", "Lin Fld FmFld ToFld Fny", L, Fld, FmFld, ToFld, A.Fny: Exit Sub
    Dim IInfDtOfFld%, IxFm%, IxTo%
        IInfDtOfFld = IxzAy(A.Fny, Fld)
        IxFm = IxzAy(A.Fny, FmFld)
        IxTo = IxzAy(A.Fny, ToFld)
    If IsBet(IInfDtOfFld, IxFm, IxTo) Then
        Inf CSub, "Fld cannot between FmFld ToFld", "Fld FmFld ToFmd IInfDtOfFld IxFm IxTo Fny", Fld, FmFld, ToFld, IxFm, IxTo, A.Fny: Exit Sub
    End If
    If IxFm > IxTo Then Inf CSub, "FmFld should be in front to ToFld", "FmFld ToFld FmIx EIx Fny", FmFld, ToFld, IxFm, IxTo, A.Fny: Exit Sub
WRg(Fld).Formula = FmtQQ("=Sum([?]:[?])", FmFld, ToFld)
End Sub
'Cor----------------------------------------------------------
Private Sub WFmtCor(L$)
Dim T1$, Fny$(): WAsgT1zFny T1, Fny, L
Dim C&: C = Colr(T1)
Dim F
For Each F In Itr(Fny)
    WRg(F).Color = C
Next
End Sub

'Fml----------------------------------------------------------
Private Sub WFmtFml(L$)
Dim Fld$, Fml$: AsgTRst L, Fld, Fml
If Not HasEle(A.Fny, Fld) Then Inf CSub, "Fld not in Fny", "Fld-not-in-Fny Fml-Lin", Fld, L: Exit Sub
WRg(Fld).Formula = Fml
End Sub

'Fmt----------------------------------------------------------
Private Sub WFmtFmt(L$)
Dim Fmt$, Fny$(): WAsgT1zFny Fmt, Fny, L
Dim F
For Each F In Itr(Fny)
    WRg(F).NumberFormat = Fmt
Next
End Sub

'Lbl----------------------------------------------------------
'Must run after forumla & Between
Private Sub WFmtLbl(L$)
Dim Fld$, Lbl$
    AsgTRst L, Fld, Lbl
    If Not HasEle(A.Fny, Fld) Then Inf CSub, "Fld not in Fny", "Fld-with-er LblLin Fny", Fld, L, A.Fny: Exit Sub
Dim R1 As Range
Dim R2 As Range
Set R1 = WHdrCell(Fld)
Set R2 = CellAbove(R1)
SwapValzRg R1, R2
End Sub

'Lvl----------------------------------------------------------
Private Sub WFmtLvl(L$)
Dim T1$, Lvl As Byte, Fny$()
WAsgT1zFny T1, Fny, L
If Not IsBet(T1, "2", "8") And Len(T1) <> 1 Then Inf CSub, "Lvl should betwee 2 to 8", "Lvl LvlLin", Lvl, L: Exit Sub
Lvl = T1
Dim F
For Each F In Itr(Fny)
    WCol(F).OutlineLevel = Lvl
Next
End Sub

'Tot----------------------------------------------------------
Private Sub WFmtTot(L$)
Dim T1$, T As XlTotalsCalculation, Fny$(): WAsgT1zFny T1, Fny, L
Dim F
T = WTotCalczStr(T1)
For Each F In Itr(Fny)
    A.Lo.ListColumns(F).Total = T
Next
End Sub

Private Function WTotCalczStr(S) As XlTotalsCalculation
Dim O As XlTotalsCalculation
Select Case S
Case "Sum": O = xlTotalsCalculationSum
Case "Avg": O = xlTotalsCalculationAverage
Case "Cnt": O = xlTotalsCalculationCount
Case Else: Inf CSub, "Invalid TotCalcStr", "TotCalcStr Valid-TotCalcStr", S, "Sum Avg Cnt": Exit Function
End Select
WTotCalczStr = O
End Function

'Wdt----------------------------------------------------------
Private Sub WFmtWdt(L$)
Dim T1$, Fny$(): WAsgT1zFny T1, Fny, L
Dim W%: W = T1
If Not IsBet(W, 5, 200) Then Inf CSub, "Invalid Wdt (should between 5 200)", "Wdt Lin", W, L: Exit Sub
Dim F
For Each F In Itr(Fny)
    WRg(F).EntireColumn.Width = W
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

Private Sub Z_WFmtBdr()
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
    WFmtBdr L      '<=='
    Stop
    Return
End Sub

Private Sub ZZ()
Dim A As ListObject
Dim B$()
Dim C As Workbook
Dim XX
End Sub

'Fun===========================================================================
Private Sub WAsgT1Fny(OT1, OFny$(), Lin)

End Sub
Private Function WRgAy(Fny$()) As Range()
Dim F
For Each F In Itr(Fny)
    PushObj WRgAy, WRg(F)
Next
End Function
Private Function WRg(F) As Range
Set WRg = A.Lo.ListColumns(F).DataBodyRange
End Function
Private Function WCol(F) As Range
Set WCol = A.Lo.ListColumns(F).DataBodyRange.EntireColumn
End Function
Private Function WLy(T1) As String()
WLy = AywRmvT1(A.Lof, T1)
End Function
Private Function WItr(T1)
Asg WLy(T1), WItr
End Function

Private Function WHdrCell(C) As Range
Set WHdrCell = A1zRg(CellAbove(A.Lo.ListColumns(C).Range))
End Function

Sub StdFmtLozWb(A As Workbook)
Dim W As Worksheet
For Each W In A.Sheets
    StdFmtLozWs W
Next
End Sub
Sub StdFmtLo(A As ListObject)
FmtLo A, StdLof
End Sub
Sub StdFmtLozWs(A As Worksheet)
Dim Lo As ListObject
For Each Lo In A.ListObjects
    FmtLo Lo, StdLof
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

Sub AutoFitLo(A As ListObject)
Dim C As Range: Set C = LoAllEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = EntRgC(C, J)
   If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Function BdrLoAround(A As ListObject)
Dim R As Range
Set R = RgzMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgzMoreBelow(R)
BdrRgAround R
End Function

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

Sub SetLccWrp(A As ListObject, CC$, Optional Wrp As Boolean)
Dim C
For Each C In TermAy(CC)
    SetLcWrp A, C, Wrp
Next
End Sub
Sub SetLcWrp(A As ListObject, C, Optional Wrp As Boolean)
A.Lo.ListColumns(C).DataBodyRange.WrapText = Wrp
End Sub

Sub SetLccWdt(A As ListObject, CC$, W)
Dim C
For Each C In TermAy(CC)
    SetLcWdt A, C, W
Next
End Sub
Sub SetLcWdt(A As ListObject, C, W)
EntColzLc(A, C).ColumnWidth = W
End Sub
Function EntColzLc(A As ListObject, C) As Range
Set EntColzLc = A.Lo.ListColumns(C).DataBodyRange.EntireColumn
End Function
Sub SetLcTotLnk(A As ListObject, C)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.Lo.ListColumns(C).DataBodyRange
Set Ws = WszRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

