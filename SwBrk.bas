VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwBrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Ix%, Nm$, OpStr$
Private A_TermAy() As String

Friend Property Get TermAy() As String()
TermAy = A_TermAy
End Property
Friend Property Let TermAy(A$())
A_TermAy = A
End Property
Friend Function Init(Ix%, Nm$, OpStr$, TermAy$()) As SwBrk
With Me
    .Ix = Ix
    .Nm = Nm
    .OpStr = OpStr
    A_TermAy = TermAy
End With
End Function
Property Get Lin$()
Lin = Quote(Ix, "L#(*) ") & QuoteSq(JnSpc(Array(Nm, OpStr, JnSpc(A_TermAy))))
End Property

