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
Private Const CMod$ = "SwBrk."
Public Ix%, Nm$, OpStr$
Private A_TermSy() As String

Friend Property Get TermSy() As String()
TermSy = A_TermSy
End Property
Friend Property Let TermSy(A$())
A_TermSy = A
End Property
Friend Function Init(Ix%, Nm$, OpStr$, TermSy$()) As SwBrk
With Me
    .Ix = Ix
    .Nm = Nm
    .OpStr = OpStr
    A_TermSy = TermSy
End With
End Function
Property Get Lin$()
Lin = Quote(Ix, "L#(*) ") & QuoteSq(JnSpc(Array(Nm, OpStr, JnSpc(A_TermSy))))
End Property

