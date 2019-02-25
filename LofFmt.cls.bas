VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofFmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Cor&
Private F$()
Public Fmt$
Friend Function Init(Fmt$, Fny$()) As LofCor
Me.Fmt = Fmt
F = Fny
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property

