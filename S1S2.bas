VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1S2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public s1$, s2$

Friend Function Init(s1, s2) As S1S2
Me.s1 = s1
Me.s2 = s2
Set Init = Me
End Function

Property Get ToStr$()
ToStr = "S1S2(S1(" & s1 & ") S2(" & s2 & "))"
End Property
