Attribute VB_Name = "QDta_B_Dt"
Option Compare Text
Option Explicit
Private Const CMod$ = "BDt."
Type DT: DtNm As String: Fny() As String: Dy() As Variant: End Type
Type Dts: N As Long: Ay() As DT: End Type

Sub BrwDt(A As DT, Optional Fnn$)
BrwAy FmtDt(A), Dft(Fnn, A.DtNm)
End Sub

Function CsvLyzDt(A As DT) As String()
Dim Dy(): Dy = A.Dy
Push CsvLyzDt, JnComma(SyQteDbl(A.Fny))
Dim QQStr$: QQStr = CsvQQStrzDr(Dr(Dy, 0))
Dim IDr
For Each IDr In A.Dy
   PushI CsvLyzDt, FmtQQAv(QQStr, CvAv(IDr))
Next
End Function

Function CsvQQStrzDr$(Dr())
Dim O$(), I
For Each I In Dr
    If IsStr(I) Then
        PushI O, """?"""""
    Else
        PushI O, "?"
    End If
Next
CsvQQStrzDr = JnComma(O)
End Function

Sub DmpDt(A As DT)
DmpAy FmtDt(A)
End Sub

Function Dr(Dy(), R&) As Variant()
Dr = Dy(R)
End Function

Function DrszDt(A As DT) As Drs
DrszDt = Drs(A.Fny, A.Dy)
End Function

Function DT(DtNm, Fny() As String, Dy() As Variant) As DT
With DT
    .DtNm = DtNm
    .Fny = Fny
    .Dy = Dy
End With
End Function

Function DtDrpCol(A As DT, CC$, Optional DtNm$) As DT
DtDrpCol = DtzDrs(DrpCol(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtReOrd(A As DT, BySubFF$) As DT
DtReOrd = DtzDrs(ReOrdCol(DrszDt(A), BySubFF), A.DtNm)
End Function

Function DtSelCol(A As DT, CC$, Optional DtNm$) As DT
DtSelCol = DtzDrs(SelDrs(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtzDrs(A As Drs, Optional DtNm$ = "Dt") As DT
DtzDrs = DT(DtNm, A.Fny, A.Dy)
End Function

Function DtzFF(DtNm$, FF$, Dy()) As DT
DtzFF = DT(DtNm, Ny(FF), Dy)
End Function

Function DtzNmDrs(DtNm$, A As Drs) As DT
DtzNmDrs = DtzDrs(A, DtNm)
End Function

Property Get EmpDtAy() As DT()
End Property

Function IsEmpDt(A As DT) As Boolean
IsEmpDt = Si(A.Dy) = 0
End Function

Function NRowzDrs&(A As Drs)
NRowzDrs = Si(A.Dy)
End Function

Function NRowzDt&(A As DT)
NRowzDt = Si(A.Dy)
End Function

Function PushDt(O As Dts, M As DT)
With O
    ReDim Preserve .Ay(.N)
    .Ay(.N) = M
    .N = .N + 1
End With
End Function

'
