VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_LnxAy() As Lnx

Property Let LnxAy(A() As Lnx)
A_LnxAy = A
End Property

Property Get LnxAy() As Lnx()
LnxAy = A_LnxAy
End Property
