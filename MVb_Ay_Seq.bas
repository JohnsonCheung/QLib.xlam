Attribute VB_Name = "MVb_Ay_Seq"
Option Explicit


Private Function IntoSeq_FmTo(FmNum, ToNum, OAy)
Dim O&()
ReDim OAy(Abs(FmNum - ToNum))
Dim J&, I&
If ToNum > FmNum Then
    For J = FmNum To ToNum
        OAy(I) = J
        I = I + 1
    Next
Else
    For J = ToNum To FmNum Step -1
        OAy(I) = J
        I = I + 1
    Next
End If
IntoSeq_FmTo = OAy
End Function
Function CvIntAy(A) As Integer()
CvIntAy = A
End Function
Function CvLngAy(A) As Long()
CvLngAy = A
End Function
Function IntSeq_FmTo(FmNum%, ToNum%) As Integer()
IntSeq_FmTo = CvIntAy(IntoSeq_FmTo(FmNum, ToNum, EmpIntAy))
End Function

Function LngSeq_FmTo(FmNum&, ToNum&) As Long()
LngSeq_FmTo = CvLngAy(IntoSeq_FmTo(FmNum, ToNum, EmpLngAy))
End Function

Function IntSeq_0U(U%) As Integer()
IntSeq_0U = IntSeq_FmTo(0, U)
End Function

Function LngSeq_0U(U&) As Long()
LngSeq_0U = LngSeq_FmTo(0, U)
End Function

