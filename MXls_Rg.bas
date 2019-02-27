Attribute VB_Name = "MXls_Rg"
Option Explicit
Function CvRg(A) As Range
Set CvRg = A
End Function

Function RgA1(A As Range) As Range
Set RgA1 = RgRC(A, 1, 1)
End Function
Function A1(Ws As Worksheet) As Range
Set A1 = WsRC(Ws, 1, 1)
End Function
Function A1zRg(A As Range) As Range
Set A1zRg = A1(WszRg(A))
End Function
Function IsA1(A As Range) As Boolean
If A.Row <> 1 Then Exit Function
If A.Column <> 1 Then Exit Function
IsA1 = True
End Function
Function RgAdr$(A As Range)
RgAdr = "'" & WszRg(A).Name & "'!" & A.Address
End Function
Function AdrRg$(A As Range)
AdrRg = RgAdr(A)
End Function

Sub AsgRRRCCRg(A As Range, OR1, OR2, OC1, OC2)
OR1 = A.Row
OR2 = OR1 + A.Rows.Count - 1
OC1 = A.Column
OC2 = OC1 + A.Columns.Count - 1
End Sub

Sub BdrRg(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub

Sub BdrRgAround(A As Range)
BdrRgLeft A
BdrRgRight A
BdrRgTop A
BdrRgBottom A
End Sub

Sub BdrRgBottom(A As Range)
BdrRg A, xlEdgeBottom
BdrRg A, xlEdgeTop
End Sub

Sub BdrRgInner(A As Range)
BdrRg A, xlInsideHorizontal
BdrRg A, xlInsideVertical
End Sub

Sub BdrRgInside(A As Range)
BdrRgInner A
End Sub
Sub BdrRgAlign(A As Range, H As XlHAlign)
Select Case H
Case XlHAlign.xlHAlignLeft: BdrRgLeft A
Case XlHAlign.xlHAlignRight: BdrRgRight A
End Select
End Sub
Sub BdrRgLeft(A As Range)
BdrRg A, xlEdgeLeft
If A.Column > 1 Then
    BdrRg RgC(A, 0), xlEdgeRight
End If
End Sub

Sub BdrRgRight(A As Range)
BdrRg A, xlEdgeRight
If A.Column < MaxWsCol Then
    BdrRg RgC(A, A.Column + 1), xlEdgeLeft
End If
End Sub

Sub BdrRgTop(A As Range)
BdrRg A, xlEdgeTop
If A.Row > 1 Then
    BdrRg RgC(A, A.Row + 1), xlEdgeBottom
End If
End Sub

Function RgC(A As Range, C) As Range
Set RgC = RgCC(A, C, C)
End Function

Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, NRowRg(A), C2)
End Function

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function EntRgC(A As Range, C) As Range
Set EntRgC = RgC(A, C).EntireColumn
End Function

Function EntRgRR(A As Range, R1, R2) As Range
Set EntRgRR = RgRR(A, R1, R2).EntireRow
End Function

Function FstColRg(A As Range) As Range
Set FstColRg = RgC(A, 1)
End Function

Function FstRowRg(A As Range) As Range
Set FstRowRg = RgR(A, 1)
End Function

Function IsHBarRg(A As Range) As Boolean
IsHBarRg = A.Rows.Count = 1
End Function

Function IsVBarRg(A As Range) As Boolean
IsVBarRg = A.Columns.Count = 1
End Function

Function LasColRg%(A As Range)
LasColRg = A.Column + A.Columns.Count - 1
End Function

Function LasHBarRg(A As Range) As Range
Set LasHBarRg = RgR(A, NRowRg(A))
End Function

Function LasRowRg&(A As Range)
LasRowRg = A.Row + A.Rows.Count - 1
End Function

Function LasVBarRg(A As Range) As Range
Set LasVBarRg = RgC(A, NColRg(A))
End Function


Function LozSq(Sq(), At As Range, Optional LoNm$) As ListObject
Set LozSq = LozRg(RgzSq(Sq, At), LoNm)
End Function
Function LozRg(Rg As Range, Optional LoNm$) As ListObject
Dim O As ListObject: Set O = WszRg(Rg).ListObjects.Add(xlSrcRange, Rg, , xlYes)
BdrRgAround Rg
Rg.EntireColumn.AutoFit
Set LozRg = SetLoNm(O, LoNm)
End Function

Sub MgeRg(A As Range)
A.MergeCells = True
A.HorizontalAlignment = XlHAlign.xlHAlignCenter
A.VerticalAlignment = XlVAlign.xlVAlignCenter
End Sub

Function NColRg%(A As Range)
NColRg = A.Columns.Count
End Function

Function RgzMoreBelow(A As Range, Optional N% = 1)
Set RgzMoreBelow = RgRR(A, 1, A.Rows.Count + N)
End Function

Function RgzMoreTop(A As Range, Optional N = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, A.Rows.Count)
Set RgzMoreTop = O
End Function

Function NRowRg&(A As Range)
NRowRg = A.Rows.Count
End Function

Function RgR(A As Range, R)
Set RgR = RgRR(A, R, R)
End Function

Function CellBelow(Cell As Range, Optional N = 1) As Range
Set CellBelow = RgRC(Cell, 1 + N, 1)
End Function

Sub SwapValzRg(Cell1 As Range, Cell2 As Range)
Dim A: A = RgRC(Cell1, 1, 1).Value
RgRC(Cell1, 1, 1).Value = RgRC(Cell2, 1, 1).Value
RgRC(Cell2, 1, 1).Value = A
End Sub
Function CellAbove(Cell As Range, Optional Above = 1) As Range
Set CellAbove = RgRC(Cell, 1 - Above, 1)
End Function

Function CellRight(A As Range, Optional Right = 1) As Range
Set CellRight = RgRC(A, 1, 1 + Right)
End Function
Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = WszRg(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, NColRg(A))
End Function

Function RgzResz(At As Range, Sq) As Range
Set RgzResz = RgRCRC(At, 1, 1, NRowSq(Sq), NColSq(Sq))
End Function

Function SqzRg(A As Range) As Variant()
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        SqzRg = O
        Exit Function
    End If
End If
SqzRg = A.Value
End Function

Function WbzRg(A As Range) As Workbook
Set WbzRg = WbzWs(WszRg(A))
End Function
Function WszRg(A As Range) As Worksheet
Set WszRg = A.Parent
End Function

Private Sub Z_RgzMoreBelow()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = NewWs
Set R = Ws.Range("A3:B5")
Set Act = RgzMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub

Private Sub Z()
Z_RgzMoreBelow
MXls_Z_Rg:
End Sub
