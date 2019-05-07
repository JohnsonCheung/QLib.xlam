Attribute VB_Name = "QXls_Cell_Clr"
Option Explicit
Private Const CMod$ = "MXls_Cell_Clr."
Private Const Asm$ = "QXls"
Sub ClrCellBelow(Cell As Range)
RgzBelowCell(Cell).Clear
End Sub
Function RgzBelowCell(Cell As Range)
Dim Ws As Worksheet: Set Ws = WszRg(Cell)
Set RgzBelowCell = WsCRR(Ws, Cell.Column, Cell.Row, LasRno(Ws))
End Function
