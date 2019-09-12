Attribute VB_Name = "MxSeq"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSeq."


Private Function IntoSeqzFT(OInto, FmNum, ToNum)
Dim O&()
ReDim OInto(Abs(FmNum - ToNum))
Dim J&, I&
If ToNum > FmNum Then
    For J = FmNum To ToNum
        OInto(I) = J
        I = I + 1
    Next
Else
    For J = ToNum To FmNum Step -1
        OInto(I) = J
        I = I + 1
    Next
End If
IntoSeqzFT = OInto
End Function

Function CvIntAy(A) As Integer()
On Error Resume Next
CvIntAy = A
End Function

Function CvLngAy(A) As Long()
On Error Resume Next
CvLngAy = A
End Function

Function IntSeqzFT(FmNum%, ToNum%) As Integer()
IntSeqzFT = IntoSeqzFT(EmpIntAy, FmNum, ToNum)
End Function

Function LngSeqzFT(FmNum&, ToNum&) As Long()
LngSeqzFT = IntoSeqzFT(EmpLngAy, FmNum, ToNum)
End Function

Function IntSeqzU(U%) As Integer()
IntSeqzU = IntSeqzFT(0, U)
End Function

Function LngSeqzU(U&) As Long()
LngSeqzU = LngSeqzFT(0, U)
End Function
