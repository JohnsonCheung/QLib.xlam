VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Fxn$, Wsn$, T$, Bexpr$
Private A_Fxc() As LidFxc

Friend Function Init(Fxn$, Wsn$, T$, Fxc() As LidFxc, Bexpr$) As LidFx
With Me
.Fxn = Fxn
.T = T
.Wsn = Wsn
.Bexpr = Bexpr
End With
A_Fxc = Fxc
Set Init = Me
End Function

Property Get Fxc() As LidFxc()
Fxc = A_Fxc
End Property
