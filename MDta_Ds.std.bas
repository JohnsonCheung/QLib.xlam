Attribute VB_Name = "MDta_Ds"
Option Explicit
Function DsAddDt(A As Ds, T As DT) As Ds
If DsHasDt(A, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", A.DsNm, T.DtNm)
Dim N%: N = Sz(A.DtAy)
Dim Ay() As DT
    Ay = A.DtAy
ReDim Preserve Ay(N)
Set Ay(N) = T
Set DsAddDt = Ds(Ay, A.DsNm)
End Function
Function CvDs(A) As Ds
Set CvDs = A
End Function
Function DsAddDtAy(A As Ds, B() As DT) As Ds
Dim I, O As Ds
Set O = A
For Each I In B
    Set O = DsAddDt(O, CvDt(I))
Next
Set DsAddDtAy = O
End Function

Function DsDt(A As Ds, Ix%) As DT
Dim DtAy() As DT
DtAy = A.DtAy
Set DsDt = DtAy(Ix)
End Function
Function Ds(A() As DT, Optional DsNm$ = "Ds") As Ds
Dim O As New Ds
Set Ds = O.Init(A, DsNm)
End Function

Function DsHasDt(A As Ds, DtNm) As Boolean
Dim DT
For Each DT In Itr(A.DtAy)
    If CvDt(DT).DtNm = DtNm Then DsHasDt = True: Exit Function
Next
End Function

Function DsIsEmp(A As Ds) As Boolean
DsIsEmp = Sz(A.DtAy) = 0
End Function


Function DsNDt%(A As Ds)
DsNDt = Sz(A.DtAy)
End Function
