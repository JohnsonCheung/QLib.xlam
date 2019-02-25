Attribute VB_Name = "MXls_Lo_Fmt_Col"
Option Explicit
Private A As ListObject, B As Lof
Sub FmtLoz(Lo As ListObject, Lof As Lof)
Set A = Lo
Dim J%
With Lof
End With
End Sub


Private Sub FmtFmlz(C, Fml$)
RgLoC(A, C).Formula = Fml
End Sub

Private Sub FmtCorz(C, Cor&)
RgLoC(A, C).Interior.Color = Cor
End Sub

Private Sub FmtFmtz(C, Fmt$)
RgLoC(A, C).Formula = Fmt
End Sub

Private Sub FmtLvlz(C, Optional Lvl As Byte = 2)
RgLoC(A, C).EntireColumn.OutlineLevel = Lvl
End Sub

Private Sub FmtWdtz(C, W%)
RgLoC(A, C).EntireCumn.ColumnWidth = W
End Sub

Private Sub FmtAliz(C, B As XlHAlign)
RgLoC(A, C).HorizontalAlignment = B
End Sub

Private Function RgLoC(A As ListObject, C) As Range
Set RgLoC = A.ListColumns(C).DataBodyRange
End Function

Private Sub FmtBdrLeftz(C)
BdrRgLeft RgLoC(A, C)
End Sub

Private Sub FmtBdrRightz(C)
BdrRgLeft RgLoC(A, C)
End Sub

Private Sub FmtTotz(C, B As XlTotalsCalculation)
A.ListColumns(C).TotalsCalculation = B
End Sub

Private Sub FmtFml(Cy$(), Fml$)
RgLoC(A, Cy).Formula = Fml
End Sub

Private Sub FmtCor(Cy$(), Cor&)
RgLoC(A, Cy).Interior.Color = Cor
End Sub

Private Sub FmtFmt(Cy$(), Fmt$)
RgLoC(A, Cy).Formula = Fmt
End Sub

Private Sub FmtLvl(Cy$(), Optional Lvl As Byte = 2)
RgLoC(A, Cy).EntireColumn.OutlineLevel = Lvl
End Sub

Private Sub FmtWdt(Cy$(), W%)
RgLoC(A, Cy).EntireCumn.ColumnWidth = W
End Sub

Private Sub FmtAli(Cy$(), B As XlHAlign)
RgLoC(A, Cy).HorizontalAlignment = B
End Sub
Private Sub FmtBdr(Cy$(), B As XlHAlign)
Dim C
For Each C In Itr(Cy)
    BdrRgAlign RgLoC(A, C), B
Next
End Sub

Private Sub FmtBdrLeft(Cy$())
Dim C
For Each C In Itr(Cy)
    BdrRgLeft RgLoC(A, C)
Next
End Sub

Private Sub FmtBdrRight(Cy$())
BdrRgLeft RgLoC(A, Cy)
End Sub

Private Sub FmtTot(Cy$(), B As XlTotalsCalculation)
'A.ListColumns(C).TotalsCalculation = B
End Sub

