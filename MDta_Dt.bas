Attribute VB_Name = "MDta_Dt"
Option Explicit
Function DtAddAp(A As Dt, ParamArray DtAp()) As Dt()
Dim O() As Dt, Av(), I
PushObj O, A
Av = DtAp
For Each I In Av
    If Not IsDt(I) Then Thw CSub, "Given DtAp should all be Dt", "Type-I-Not-Dt", TypeName(I)
    PushObj O, CvDt(I)
Next
DtAddAp = O
End Function

Function IsDt(A) As Boolean
IsDt = TypeName(A) = "Dt"
End Function

Sub BrwDs(A As Ds, Optional Fnn$)
BrwAy FmtDs(A), Dft(Fnn, A.DsNm)
End Sub

Sub BrwDt(A As Dt, Optional Fnn$)
BrwAy FmtDt(A), Dft(Fnn, A.DtNm)
End Sub

Function CsvQQStrDr$(Dr)
Dim O$(), I
For Each I In Dr
    If IsStr(I) Then
        PushI O, """?"""""
    Else
        PushI O, "?"
    End If
Next
CsvQQStrDr = JnComma(O)
End Function
Function CsvLyDt(A As Dt) As String()
Dim Dry(): Dry = A.Dry
Push CsvLyDt, JnComma(AyQuoteDbl(A.Fny))
Dim QQStr$: 'QQStr = CsvQuotestrDr(Dry(0))
Dim Dr
For Each Dr In A.Dry
   Push CsvLyDt, FmtQQAv(QQStr, CvAv(Dr))
Next
End Function

Function DtSelCol(A As Dt, CC, Optional DtNm$) As Dt
Set DtSelCol = DtDrsDtnm(DrsSelCC(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtDrpCol(A As Dt, CC, Optional DtNm$) As Dt
Set DtDrpCol = DtDrsDtnm(DrsDrpCC(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DrszDt(A As Dt) As Drs
Set DrszDt = Drs(A.Fny, A.Dry)
End Function
Function DtzDrs(A As Drs, Optional DtNm$ = "Dt") As Dt
Set DtzDrs = Dt(DtNm, A.Fny, A.Dry)
End Function

Function NRowzDt&(A As Dt)
NRowzDt = Sz(A.Dry)
End Function
Function NRowzDrs&(A As Drs)
NRowzDrs = Sz(A.Dry)
End Function
Sub DmpDt(A As Dt)
DmpAy FmtDt(A)
End Sub
Property Get EmpDtAy() As Dt()
End Property

Function IsEmpDt(A As Dt) As Boolean
IsEmpDt = Sz(A.Dry) = 0
End Function

Function DtReOrd(A As Dt, BySubFF) As Dt
Set DtReOrd = DtDrsDtnm(DrsReOrdBy(DrszDt(A), BySubFF), A.DtNm)
End Function
Function Dt(DtNm, Fny0, Dry()) As Dt
Dim O As New Dt
Set Dt = O.Init(DtNm, Fny0, Dry)
End Function

Function CvDt(A) As Dt
Set CvDt = A
End Function

Function DtAy(ParamArray Ap()) As Dt()
Dim Av(): Av = Ap
DtAy = IntozAy(DtAy, Av)
End Function

Private Sub ZZ_DtAy()
Dim A() As Dt
A = DtAy(SampDt1, SampDt2)
Stop
End Sub


