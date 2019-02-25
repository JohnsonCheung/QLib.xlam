VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofLvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private F$()
Public Lvl As Byte
Friend Function Init(Lvl As Byte, Fny$()) As LofAli
F = Fny
Me.Lvl = Lvl
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property


