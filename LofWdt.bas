VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofWdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Wdt%
Private F$()
Public Fmt$
Friend Function Init(Wdt%, Fny$()) As LofCor
Me.Wdt = Wdt
F = Fny
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property


