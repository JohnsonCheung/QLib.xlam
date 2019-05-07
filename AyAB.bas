VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AyAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "AyAB."
Private Type A
    A As Variant
    B As Variant
End Type
Private X As A
Friend Function Init(A, B) As Ayab
ThwIfNotAy A, CSub
ThwIfNotAy B, CSub
Set Init = Me
End Function
Property Get A()
A = X.A
End Property
Property Get B()
B = X.B
End Property
