Attribute VB_Name = "QXls_Lo_FmtLoTit"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Lo_Fmt_Tit."
Private Const Asm$ = "QXls"

Sub SetLoTit(A As ListObject, TitLy$())
Dim Sq(), R As Range
    Sq = TitSq(TitLy, FnyzLo(A)): If Si(Sq) = 0 Then Exit Sub
    Set R = TitAt(A, UBound(Sq(), 1))
Set R = RgzSq(Sq(), R)
MgeRg R
BdrRgInner R
BdrRgAround R
End Sub

Private Sub MgeRg(TitRg As Range)
Dim J%
For J = 1 To TitRg.Rows.Count
    MgeRgH RgR(TitRg, J)
Next
For J = 1 To TitRg.Columns.Count
    MgeRgV RgC(TitRg, J)
Next
End Sub

Private Sub MgeRgH(TitRg As Range)
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

Private Sub MgeRgV(A As Range)
Dim J%
For J = A.Rows.Count To 2 Step -1
    MgeCellAbove RgRC(A, J, 1)
Next
End Sub

Private Function TitAt(Lo As ListObject, NTitRow%) As Range
Set TitAt = RgRC(Lo.DataBodyRange, 0 - NTitRow, 1)
End Function

Private Function TitSq(TitLy$(), LoFny$()) As Variant()
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
TitSq = Transpose(SqzDry(Col))
End Function

Private Sub Z_TitSq()
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
Ept = TitSq(TitLy, Fny)
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
BrwSq TitSq(TitLy, Fny)

Exit Sub
Tst:
    Act = TitSq(TitLy, Fny)
    Ass IsEqSq(Act, Ept)
    Return
End Sub

Private Sub ZZ()
QXls_Lo_FmtLoTit:
Z_TitSq
End Sub
