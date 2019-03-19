Attribute VB_Name = "MXls_RgVBar"
Option Explicit
Sub Vbar_MgeBottomEmpCell(A As Range)
Ass IsVbarRg(A)
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

Function VbarAy(A As Range) As Variant()
Ass IsVbarRg(A)
'VbarAy = Sq_Col(RgzSq(A), 1)
End Function

Function VbarIntAy(A As Range) As Integer()
'VbarIntAy = AyIntAy(VbarAy(A))
End Function

Function VbarSy(A As Range) As String()
VbarSy = SyzAy(VbarAy(A))
End Function
