Attribute VB_Name = "MTp_Tp_Gp"
Option Explicit
Function GpzLy(Ly$()) As Gp
Set GpzLy = Gp(LnxAy(Ly))
End Function

Function CvGp(A) As Gp
Set CvGp = A
End Function

Function LyzGp(A As Gp) As String()
LyzGp = LyzLnxAy(A.LnxAy)
End Function

Function Gp(A() As Lnx) As Gp
Set Gp = New Gp
With Gp
    .LnxAy = A
End With
End Function

Function GpAyzBlkTy(A() As Blk, BlkTy$) As Gp()
Dim J%
For J = 0 To UB(A)
    With A(J)
        If .BlkTy = BlkTy Then
            PushObj GpAyzBlkTy, A(J).Gp
        End If
    End With
Next
End Function


