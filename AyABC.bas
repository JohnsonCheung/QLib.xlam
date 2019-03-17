VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AyABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type A
    A As Variant
    B As Variant
    C As Variant
End Type
Private X As A
Friend Function Init(A, B, C) As AyABC
ThwNotAy A, CSub
ThwNotAy B, CSub
ThwNotAy C, CSub
With X
    .A = A
    .B = B
    .C = C
End With
Set Init = Me
End Function
Property Get A()
A = X.A
End Property
Property Get B()
B = X.B
End Property
Property Get C()
C = X.C
End Property
