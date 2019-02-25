Attribute VB_Name = "MXls_Lo_Fmt"
Option Explicit
Const CMod$ = "MXls_Lo_Fmt."
Private A As ListObject, B As Lof
Function FmtLo(Lo As ListObject, Fmtr$()) As ListObject
Set A = Lo
Const CSub$ = CMod & "FmtLo"
Dim D As Dictionary
Set D = DefDic(Fmtr, LofKK)
ThwErMsg LofEr(D), CSub, "There is error in Fmtr", "Fmtr LoNm", Fmtr, A.Name
Set B = Lof(D)
FmtAli
FmtBdr
FmtBet
FmtCor
FmtFml
FmtFmt
FmtLbl
FmtLvl
FmtTit
FmtTot
FmtWdt
Set FmtLo = A
End Function

Sub FmtLozWb(Wb As Workbook, LoNmToFmtrVblDic As Dictionary)
Dim I, L As ListObject
For Each I In Itr(LoAy(Wb))
    Set L = CvLo(I)
    FmtLoz L, Lof(CvLo(L))
Next
End Sub

Private Sub FmtAli()

End Sub

Private Sub FmtWdt()
End Sub
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

Private Sub Z_Bdr()
Dim mBdr() As LofBdr
'--
Erase mBdr
PushI mBdr, "Left A B C"
PushI mBdr, "Left D E F"
PushI mBdr, "Right A B C"
PushI mBdr, "Center A B C"
GoSub Tst
Tst:
    
'    Bdr      '<=='
    Return
End Sub

Private Sub ZZ()
Dim A As ListObject
Dim B$()
Dim C As Workbook
Dim XX
End Sub

Private Sub Z()
End Sub


Sub Set1LcBet()
Dim L, C$, X$, Y$
'Dim Tot() As LofBet: X = B.Tot
'For Each I In Itr(Tot)
    Asg2TRst L, C, X, Y
'    A_Lo.ListColumns(C).DataBodyRange.Formula = FmtQQ("=Sum([?]:[?])", X, Y)
'Next
End Sub

Sub Set1LcFldLikssCalc(FldLikss$, B As XlTotalsCalculation)
Dim F
'For Each F In A_Fny
'    If StrLikss(F, FldLikss) Then Set1LcTot A_Lo, F, B
'Next
End Sub

Sub Set1LcTot()
'Set1LcFldLikssCalc FstEleRmvT1(Tot, "Sum"), xlTotalsCalculationSum
'Set1LcFldLikssCalc FstEleRmvT1(Tot, "Cnt"), xlTotalsCalculationCount
'Set1LcFldLikssCalc FstEleRmvT1(Tot, "Avg"), xlTotalsCalculationAverage
End Sub

Sub SetLcLvl(A As ListObject, C, Lvl As Byte)

End Sub
Sub SetLcWdt(A As ListObject, C, Wdt%)

End Sub
Sub SetLcWdtzCC(A As ListObject, CC, Wdt%)

End Sub
Private Sub FmtBdr()
End Sub
Private Sub FmtBet()
End Sub
Private Sub FmtCor()
End Sub
Private Sub FmtFml()
End Sub
Private Sub FmtFmt()
End Sub
Private Sub FmtLbl()
End Sub
Private Sub FmtLvl()
End Sub
Private Sub FmtTit()
End Sub
Private Sub FmtTot()
End Sub

