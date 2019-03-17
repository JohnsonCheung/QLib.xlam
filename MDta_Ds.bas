Attribute VB_Name = "MDta_Ds"
Option Explicit
Function DsAddDt(A As Ds, T As Dt) As Ds
If DsHasDt(A, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", A.DsNm, T.DtNm)
Dim N%: N = Si(A.DtAy)
Dim Ay() As Dt
    Ay = A.DtAy
ReDim Preserve Ay(N)
Set Ay(N) = T
Set DsAddDt = Ds(Ay, A.DsNm)
End Function
Function CvDs(A) As Ds
Set CvDs = A
End Function
Function DsAddDtAy(A As Ds, B() As Dt) As Ds
Dim I, O As Ds
Set O = A
For Each I In B
    Set O = DsAddDt(O, CvDt(I))
Next
Set DsAddDtAy = O
End Function

Function DsDt(A As Ds, Ix%) As Dt
Dim DtAy() As Dt
DtAy = A.DtAy
Set DsDt = DtAy(Ix)
End Function
Function Ds(A() As Dt, Optional DsNm$ = "Ds") As Ds
Dim O As New Ds
Set Ds = O.Init(A, DsNm)
End Function

Function DsHasDt(A As Ds, DtNm) As Boolean
Dim Dt
For Each Dt In Itr(A.DtAy)
    If CvDt(Dt).DtNm = DtNm Then DsHasDt = True: Exit Function
Next
End Function

Function DsIsEmp(A As Ds) As Boolean
DsIsEmp = Si(A.DtAy) = 0
End Function


Function DsNDt%(A As Ds)
DsNDt = Si(A.DtAy)
End Function
