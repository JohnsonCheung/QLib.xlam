VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Apn$, CpyInpWsToOupFx As Boolean
Private A_Fx() As LidFx
Private A_Fb() As LidFb
Private A_Fil() As LidFil
Friend Function Init(Apn$, Fil() As LidFil, Fx() As LidFx, Fb() As LidFb, Optional CpyInpWsToOupFx As Boolean) As LidPm
Me.Apn = Apn
Me.CpyInpWsToOupFx = CpyInpWsToOupFx
A_Fil = Fil
A_Fx = Fx
A_Fb = Fb
Set Init = Me
End Function
Property Get AppFb$()
AppFb = AppHom & Apn & ".accdb"
End Property

Property Get Fil() As LidFil()
Fil = A_Fil
End Property
Property Get Fx() As LidFx()
Fx = A_Fx
End Property
Property Get Fb() As LidFb()
Fb = A_Fb
End Property

