Attribute VB_Name = "QVb_Ay_Seq"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Seq."
Private Const Asm$ = "QVb"


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
On Error GoTo X
If IsLngAy(A) Then CvLngAy = A: Exit Function
Dim I: For Each I In A
    PushI CvLngAy, I
Next
Exit Function
X:
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

