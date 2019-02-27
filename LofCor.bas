VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Cor&
Private F$()
Friend Function Init(Cor, Fny$()) As LofCor
Me.Cor = Cor
F = Fny
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property
