Attribute VB_Name = "QDta_Dta_Dt"
Option Compare Text
Option Explicit
Private Const CMod$ = "BDt."
Type Dt: DtNm As String: Fny() As String: Dry() As Variant: End Type
Type Dts: N As Long: Ay() As Dt: End Type
Function PushDt(O As Dts, M As Dt)
With O
    ReDim Preserve .Ay(.N)
    .Ay(.N) = M
    .N = .N + 1
End With
End Function

Sub BrwDt(A As Dt, Optional Fnn$)
BrwAy FmtDt(A), Dft(Fnn, A.DtNm)
End Sub

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
Function Dr(Dry(), R&) As Variant()
Dr = Dry(R)
End Function
Function CsvLyzDt(A As Dt) As String()
Dim Dry(): Dry = A.Dry
Push CsvLyzDt, JnComma(SyQuoteDbl(A.Fny))
Dim QQStr$: QQStr = CsvQQStrzDr(Dr(Dry, 0))
Dim IDr
For Each IDr In A.Dry
   PushI CsvLyzDt, FmtQQAv(QQStr, CvAv(IDr))
Next
End Function

Function DtSelCol(A As Dt, CC$, Optional DtNm$) As Dt
DtSelCol = DtzDrs(SelDrsCC(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtDrpCol(A As Dt, CC$, Optional DtNm$) As Dt
DtDrpCol = DtzDrs(DrpCny(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DrszDt(A As Dt) As Drs
DrszDt = Drs(A.Fny, A.Dry)
End Function

Function DtzDrs(A As Drs, Optional DtNm$ = "Dt") As Dt
DtzDrs = Dt(DtNm, A.Fny, A.Dry)
End Function

Function NRowzDt&(A As Dt)
NRowzDt = Si(A.Dry)
End Function
Function NRowzDrs&(A As Drs)
NRowzDrs = Si(A.Dry)
End Function
Sub DmpDt(A As Dt)
DmpAy FmtDt(A)
End Sub
Property Get EmpDtAy() As Dt()
End Property

Function IsEmpDt(A As Dt) As Boolean
IsEmpDt = Si(A.Dry) = 0
End Function

Function DtReOrd(A As Dt, BySubFF$) As Dt
DtReOrd = DtzDrs(ReOrdCol(DrszDt(A), BySubFF), A.DtNm)
End Function

Function DtzFF(DtNm$, FF$, Dry()) As Dt
DtzFF = Dt(DtNm, Ny(FF), Dry)
End Function

Function Dt(DtNm, Fny() As String, Dry() As Variant) As Dt
With Dt
    .DtNm = DtNm
    .Fny = Fny
    .Dry = Dry
End With
End Function


