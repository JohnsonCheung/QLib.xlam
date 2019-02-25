Attribute VB_Name = "MDao_Lid_PmFmt"
Option Explicit
Private B As LidPm
Function FmtLidPm(A As LidPm) As String()
Set B = A
PushI FmtLidPm, ApnLin
PushIAy FmtLidPm, Fil
PushIAy FmtLidPm, Fb
PushIAy FmtLidPm, Fx
End Function
Private Property Get ApnLin$()
ApnLin = "LidPm " & B.Apn
End Property
Private Function Fil() As LidFil()

End Function
Private Property Get Fx() As String()
Dim J%, Ay() As LidFx
Ay = B.Fx
For J = 0 To UB(Ay)
    PushI Fx, FxLin(Ay(J))
Next
End Property
Private Function FxLin$(A As LidFx)

End Function
Private Property Get Fb() As String()

End Property


