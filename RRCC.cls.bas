VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RRCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public R1&, R2&, C1%, C2%
Friend Function Init(R1&, R2&, C1%, C2%) As RRCC
With Me
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
Set Init = Me
End Function
Property Get IsEmp() As Boolean
IsEmp = True
With Me
   If .R1 <= 0 Then Exit Property
   If .R2 <= 0 Then Exit Property
   If .R1 > .R2 Then Exit Property
End With
IsEmp = False
End Property

