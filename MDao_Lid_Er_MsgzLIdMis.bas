Attribute VB_Name = "MDao_Lid_Er_MsgzLIdMis"
Option Explicit

Function MsgzLidMis(A As LidMis) As String()
MsgzLidMis = AyAddAp(MsgzMisFfnAset(A.Ffn), Tbl(A.Tbl), Col(A.Col), Ty(A.Ty))
End Function

Private Function Tbl(A() As LidMisTbl) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy Tbl, A(J).MisMsg
Next
End Function

Private Function Col(A() As LidMisCol) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy Col, A(J).MisMsg
Next
End Function

Private Property Get Ty(A() As LidMisTy) As String()
Dim J%
For J = 0 To UB(A)
    PushIAy Ty, A(J).MisMsg
Next
End Property

Private Sub Z_MsgzLidMis()
Dim LidMis As LidMis
Set LidMis = SampLidMis
GoSub Tst
Exit Sub
Tst:
    Act = MsgzLidMis(LidMis)
    D Act
    Stop
    If Not IsEqAy(Act, Ept) Then Stop
    Return
End Sub


