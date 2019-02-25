Attribute VB_Name = "MXls_RgVBar"
Option Explicit
Sub VBar_MgeBottomEmpCell(A As Range)
Ass IsVBarRg(A)
Dim R2: R2 = A.Rows.Count
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(A, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(A, 1, R1, R2)
R.Merge
R.VerticalAlignment = XlVAlign.xlVAlignTop
End Sub

Function VBarAy(A As Range) As Variant()
Ass IsVBarRg(A)
'VBarAy = Sq_Col(RgzSq(A), 1)
End Function

Function VBarIntAy(A As Range) As Integer()
'VBarIntAy = AyIntAy(VBarAy(A))
End Function

Function VBarSy(A As Range) As String()
VBarSy = SyzAy(VBarAy(A))
End Function
