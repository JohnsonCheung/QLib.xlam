VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidMisTy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Fx As String
Public Fxn As String
Public Wsn As String
Private A_Tyc() As LidMisTyc
Friend Function Init(Fx, Fxn, Wsn, Tyc() As LidMisTyc) As LidMisTy
With Me
    .Fx = Fx
    .Fxn = Fxn
    .Wsn = Wsn
End With
A_Tyc = Tyc
Set Init = Me
End Function
Property Get TycAy() As LidMisTyc()
TycAy = A_Tyc
End Property
Property Get MisMsg() As String()
MisMsg = MisMsgTyOneFx(MisMsgColMsgAy(A_Tyc))
End Property

Private Function MisMsgTyOneFx(ColMsg$()) As String()
Dim M$
    Select Case Sz(ColMsg)
    Case 0: Exit Function
    Case 1: M = "There is one column having unexpected column type"
    Case Else: M = FmtQQ("There are ? columns having unexpected column type", Sz(ColMsg))
    End Select
Dim Fxn$
Dim NN$: NN = FmtQQ("[? excel file] Worksheet unexpected", Fxn)
MisMsgTyOneFx = LyzMsgNap(M, NN, Fx, Wsn, ColMsg)
End Function

Private Function MisMsgColMsgAy(A() As LidMisTyc) As String()
Dim J%, O$()
For J = 0 To UBound(A)
    PushNonBlankStr O, A(J).MisMsg
Next
MisMsgColMsgAy = FmtAyzSepSS(O, "has [it should]")
End Function


