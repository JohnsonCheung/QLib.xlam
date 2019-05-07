Attribute VB_Name = "QXls_RgCell"
Option Explicit
Private Const CMod$ = "MXls_RgCell."
Private Const Asm$ = "QXls"
Function VyrzAt(At As Range) As Variant()
VyrzAt = VyrzSqr(SqzRg(HRgWiVal(A1(At))))
End Function
Function VyczAt(At As Range) As Variant()
VyczAt = VyczSqc(SqzRg(VRgWiVal(A1(At))))
End Function

Sub CellClrDown(A As Range)
VbarRgAt(A, AtLeastOneCell:=True).Clear
End Sub

Sub CellFillSeqDown(A As Range, FmNum&, ToNum&)
'AyRgV LngSeq(FmNum, ToNum), A
End Sub

Function IsCellInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Function
R2 = R1 + Rg.Rows.Count
If R > R2 Then Exit Function
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Function
C2 = C1 + Rg.Columns.Count
If C > C2 Then Exit Function
IsCellInRg = True
End Function

Function IsCellInRgAp(Cell As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
'IsCellInRgAp = IsCellInRgAv(A, Av)
End Function

Function IsCellInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If IsCellInRg(A, CvRg(V)) Then IsCellInRgAv = True: Exit Function
Next
End Function

Sub MgeCellAbove(Cell As Range)
'If Not IsEmpty(A.Value) Then Exit Sub
If Cell.MergeCells Then Exit Sub
If Cell.Row = 1 Then Exit Sub
If RgRC(Cell, 0, 1).MergeCells Then Exit Sub
MgeRg RgCRR(Cell, 1, 0, 1)
End Sub

Function VbarRgAt(At As Range, Optional AtLeastOneCell As Boolean) As Range
If IsEmpty(At.Value) Then
    If AtLeastOneCell Then
        Set VbarRgAt = A1zRg(At)
    End If
    Exit Function
End If
Dim R1&: R1 = At.Row
Dim R2&
    If IsEmpty(RgRC(At, 2, 1)) Then
        R2 = At.Row
    Else
        R2 = At.End(xlDown).Row
    End If
Dim C%: C = At.Column
Set VbarRgAt = WsCRR(WszRg(At), C, R1, R2)
End Function
