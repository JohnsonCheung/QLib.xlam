Attribute VB_Name = "MDta_Dt"
Option Explicit
Function DtAddAp(A As DT, ParamArray DtAp()) As DT()
Dim O() As DT, Av(), I
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

Sub BrwDt(A As DT, Optional Fnn$)
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
Function CsvLyDt(A As DT) As String()
Dim Dry(): Dry = A.Dry
Push CsvLyDt, JnComma(AyQuoteDbl(A.Fny))
Dim QQStr$: 'QQStr = CsvQuotestrDr(Dry(0))
Dim Dr
For Each Dr In A.Dry
   Push CsvLyDt, FmtQQAv(QQStr, CvAv(Dr))
Next
End Function

Function DtSelCol(A As DT, CC, Optional DtNm$) As DT
Set DtSelCol = DtDrsDtnm(DrsSelCC(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DtDrpCol(A As DT, CC, Optional DtNm$) As DT
Set DtDrpCol = DtDrsDtnm(DrsDrpCC(DrszDt(A), CC), Dft(DtNm, A.DtNm))
End Function

Function DrszDt(A As DT) As DRs
Set DrszDt = DRs(A.Fny, A.Dry)
End Function
Function DtzDrs(A As DRs, Optional DtNm$ = "Dt") As DT
Set DtzDrs = DT(DtNm, A.Fny, A.Dry)
End Function

Function NRowzDt&(A As DT)
NRowzDt = Sz(A.Dry)
End Function
Function NRowzDrs&(A As DRs)
NRowzDrs = Sz(A.Dry)
End Function
Sub DmpDt(A As DT)
DmpAy FmtDt(A)
End Sub
Property Get EmpDtAy() As DT()
End Property

Function IsEmpDt(A As DT) As Boolean
IsEmpDt = Sz(A.Dry) = 0
End Function

Function DtReOrd(A As DT, BySubFF) As DT
Set DtReOrd = DtDrsDtnm(DrsReOrdBy(DrszDt(A), BySubFF), A.DtNm)
End Function
Function DT(DtNm, Fny0, Dry()) As DT
Dim O As New DT
Set DT = O.Init(DtNm, Fny0, Dry)
End Function

Function CvDt(A) As DT
Set CvDt = A
End Function

Function DtAy(ParamArray Ap()) As DT()
Dim Av(): Av = Ap
DtAy = IntozAy(DtAy, Av)
End Function

Private Sub ZZ_DtAy()
Dim A() As DT
A = DtAy(SampDt1, SampDt2)
Stop
End Sub


