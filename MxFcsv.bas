Attribute VB_Name = "MxFcsv"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxFcsv."

Function DrszFcsv(Fcsv$) As Drs
Dim Ly$(): Ly = LyzFt(Fcsv)
Dim Fny$(): Fny = DrzCsvLin(Ly(0))
Dim Dy()
    Dim J&: For J = 1 To UB(Ly)
        PushI Dy, SplitComma(Ly(J))
    Next
DrszFcsv = Drs(Fny, Dy)
End Function