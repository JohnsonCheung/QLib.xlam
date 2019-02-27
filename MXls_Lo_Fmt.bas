Attribute VB_Name = "MXls_Lo_Fmt"
Option Explicit
Const CMod$ = "MXls_Lo_Fmt."
Private A As ListObject, B$(), Fny$()
Function FmtLo(Lo As ListObject, Fmtr$()) As ListObject
Dim L
Set A = Lo
Fny = FnyzLo(Lo)
B = Fmtr
For Each L In WItr("Ali"): WFmtAli L: Next
For Each L In WItr("Bdr"): WFmtBdr L: Next
For Each L In WItr("Bet"): WFmtBet L: Next
For Each L In WItr("Cor"): WFmtCor L: Next
For Each L In WItr("Fml"): WFmtFml L: Next
For Each L In WItr("Fmt"): WFmtFmt L: Next
For Each L In WItr("Lvl"): WFmtLvl L: Next
For Each L In WItr("Tot"): WFmtTot L: Next
For Each L In WItr("Wdt"): WFmtWdt L: Next
SetLoTit Lo, WLy("Tit")
SetLoNm Lo, T2(FstEleT1(B, "Nm"))
For Each L In WItr("Lbl"): WFmtLbl L: Next ' Must run Last
Set FmtLo = A
End Function

'Ali -----------------------------------------------------------
Private Sub WFmtAli(L)
Dim T1$, FldLikAy$(): AsgT1FldLikAy T1, FldLikAy, L
Dim H As XlHAlign: H = HAlign(T1)
Dim F
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        WRg(F).HorizontalAlignment = H
    End If
Next
End Sub
Private Function HAlign(S) As XlHAlign
Select Case S
Case "Left": HAlign = xlHAlignLeft
Case "Right": HAlign = xlHAlignRight
Case "Center": HAlign = xlHAlignCenter
Case Else: Thw CSub, "Invalid HAlignStr", "HAlignStr", S
End Select
End Function
'Bdr -----------------------------------------------------------
Private Sub WFmtBdr(L)
Dim T1$, FldLikAy$(): AsgT1FldLikAy T1, FldLikAy, L
Dim IsLeft As Boolean, IsRight As Boolean
    Select Case T1
    Case "Left"
    Case "Right"
    Case "Both"
    Case Else: Thw CSub, "Invalid Bdr", "Invalid-Bdr-Str Valid-Bdr Lin", T1, "Left Right Both", L
    End Select
Dim F
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        If IsLeft Then BdrRg WRg(F), xlEdgeLeft
        If IsRight Then BdrRg WRg(F), xlEdgeRight
    End If
Next
End Sub
'Bet----------------------------------------------------------
Private Sub WFmtBet(L)
Dim Fld$, FmFld$, ToFld$
    Asg2TRst L, Fld, FmFld, ToFld
    Dim IsEr As Boolean
    If Not HasEle(Fny, Fld) Then IsEr = True
    If Not HasEle(Fny, FmFld) Then IsEr = True
    If Not HasEle(Fny, ToFld) Then IsEr = True
    If IsEr Then Thw CSub, "Bet-Lin error, the Fld,FmFld,ToFld not all in Fny", "Lin Fld FmFld ToFld Fny", L, Fld, FmFld, ToFld, Fny
    Dim IxFld%, IxFm%, IxTo%
        IxFld = IxzAy(Fny, Fld)
        IxFm = IxzAy(Fny, FmFld)
        IxTo = IxzAy(Fny, ToFld)
    If IsBet(IxFld, IxFm, IxTo) Then
        Thw CSub, "Fld cannot between FmFld ToFld", "Fld FmFld ToFmd IxFld IxFm IxTo Fny", Fld, FmFld, ToFld, IxFm, IxTo, Fny
    End If
    If IxFm > IxTo Then Thw CSub, "FmFld should be in front to ToFld", "FmFld ToFld FmIx ToIx Fny", FmFld, ToFld, IxFm, IxTo, Fny
WRg(Fld).Formula = FmtQQ("=Sum([?]:[?])", FmFld, ToFld)
End Sub
'Cor----------------------------------------------------------
Private Sub WFmtCor(L)
Dim T1$, FldLikAy$(): AsgT1FldLikAy T1, FldLikAy, L
Dim C&: C = Colr(T1)
Dim F
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        WRg(F).Color = C
    End If
Next
End Sub

'Fml----------------------------------------------------------
Private Sub WFmtFml(L)
Dim Fld$, Fml$: AsgTRst L, Fld, Fml
If Not HasEle(Fny, Fld) Then Thw CSub, "Fld not in Fny", "Fld-not-in-Fny Fml-Lin", Fld, L
WRg(Fld).Formula = Fml
End Sub

'Fmt----------------------------------------------------------
Private Sub WFmtFmt(L)
Dim Fmt$, FldLikAy$(): AsgT1FldLikAy Fmt, FldLikAy, L
Dim F
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        WRg(F).NumberFormat = Fmt
    End If
Next
End Sub

'Lbl----------------------------------------------------------
'Must run after forumla & Between
Private Sub WFmtLbl(L)
Dim Fld$, Lbl$
    AsgTRst L, Fld, Lbl
    If Not HasEle(Fny, Fld) Then Thw CSub, "Fld not in Fny", "Fld-with-er LblLin Fny", Fld, L, Fny
Dim R1 As Range
Dim R2 As Range
Set R1 = WHdrCell(Fld)
Set R2 = CellAbove(R1)
SwapValzRg R1, R2
End Sub

'Lvl----------------------------------------------------------
Private Sub WFmtLvl(L)
Dim T1$, Lvl As Byte, FldLikAy$()
AsgT1FldLikAy T1, FldLikAy, L
If Not IsBet(T1, "2", "8") And Len(T1) <> 1 Then Thw CSub, "Lvl should betwee 2 to 8", "Lvl LvlLin", Lvl, L
Lvl = T1
Dim F
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        WCol(F).OutlineLevel = Lvl
    End If
Next
End Sub

'Tot----------------------------------------------------------
Private Sub WFmtTot(L)
Dim T1$, T As XlTotalsCalculation, FldLikAy$(): AsgT1FldLikAy T1, FldLikAy, L
Dim F
T = WTotCalczStr(T1)
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        A.ListColumns(F).Total = T
    End If
Next
End Sub

Private Function WTotCalczStr(S$) As XlTotalsCalculation
Dim O As XlTotalsCalculation
Select Case S
Case "Sum": O = xlTotalsCalculationSum
Case "Avg": O = xlTotalsCalculationAverage
Case "Cnt": O = xlTotalsCalculationCount
Case Else: Thw CSub, "Invalid TotCalcStr", "TotCalcStr Valid-TotCalcStr", S, "Sum Avg Cnt"
End Select
WTotCalczStr = O
End Function

'Wdt----------------------------------------------------------
Private Sub WFmtWdt(L)
Dim T1$, FldLikAy$(): AsgT1FldLikAy T1, FldLikAy, L
Dim W%: W = T1
If Not IsBet(W, 5, 200) Then Thw CSub, "Invalid Wdt (should between 5 200)", "Wdt Lin", W, L
Dim F
For Each F In Fny
    If HitLikAy(F, FldLikAy) Then
        WRg(F).EntireColumn.Width = W
    End If
Next
End Sub
'Tst-------------------------------------------------------------
Private Sub Z_FmtLo()
Dim Lo As ListObject, Fmtr() As String 'Lofr
'------------
Set Lo = SampLo
Fmtr = SampLoFmtr
GoSub Tst
Exit Sub
Tst:
    FmtLo Lo, Fmtr
    Return
End Sub

Private Sub Z_WFmtBdr()
Dim L, Lo As ListObject
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
Private Function WRg(F) As Range
Set WRg = A.ListColumns(F).DataBodyRange
End Function
Private Function WCol(F) As Range
Set WCol = A.ListColumns(F).DataBodyRange.EntireColumn
End Function
Private Function WLy(T1$) As String()
WLy = ItrzRmvT1(B, T1)
End Function
Private Function WItr(T1$)
Asg WLy(T1), WItr
End Function

Private Function WHdrCell(C) As Range
Set WHdrCell = A1zRg(CellAbove(A.ListColumns(C).Range))
End Function
