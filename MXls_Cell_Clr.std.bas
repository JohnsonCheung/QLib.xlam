Attribute VB_Name = "MXls_Cell_Clr"
Option Explicit
Sub ClrCellBelow(Cell As Range)
RgzBelowCell(Cell).Clear
End Sub
Function RgzBelowCell(Cell As Range)
Dim Ws As Worksheet: Set Ws = WszRg(Cell)
Set RgzBelowCell = WsCRR(Ws, Cell.Column, Cell.Row, LasRno(Ws))
End Function
