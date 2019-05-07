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
Private Const CMod$ = "Gp."
Private A() As Lnx
Friend Function Init(LnxAy() As Lnx) As Gp
A = LnxAy
Set Init = Me
End Function

Property Get LnxAy() As Lnx()
LnxAy = A
End Property
