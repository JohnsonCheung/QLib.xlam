VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_FmIx&, A_ToIx&
Property Get FmIx&()
FmIx = A_FmIx
End Property
Property Get ToIx&()
ToIx = A_ToIx
End Property
Friend Function Init(FmIx, ToIx) As FTIx
If Not (FmIx = -1 And ToIx = -2) Then ' This is known as EmpFTIx
    If FmIx < 0 Then Stop
    If ToIx < 0 Then Stop
    If FmIx > ToIx Then Stop
End If
A_FmIx = FmIx
A_ToIx = ToIx
Set Init = Me
End Function
Property Get IsEmp() As Boolean
IsEmp = Cnt <= 0
End Property
Property Get Cnt&()
Cnt = A_ToIx - A_FmIx + 1
End Property
Property Get FmNo&()
FmNo = A_FmIx + 1
End Property
Property Get ToNo&()
ToNo = A_ToIx + 1
End Property
Property Get IsVdt() As Boolean
If A_FmIx < 0 Then Exit Property
If A_ToIx < 0 Then Exit Property
If A_FmIx > A_ToIx Then Exit Property
IsVdt = True
End Property

