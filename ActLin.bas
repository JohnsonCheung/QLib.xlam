VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const A$ = "A"
Public Act As eActLin, Lin$, Lno&
Friend Function Init(Act As eActLin, Lin$, Lno&) As ActLin
With Me
    .Act = Act
    .Lin = Lin
    .Lno = Lno
End With
Set Init = Me
End Function
Property Get Ix&()
Ix = Lno - 1
End Property
Property Get ActStr$()
ActStr = IIf(Act = eeDltLin, "Dlt", "Ins")
End Property
Function ToStr$()
ToStr = ActStr & ":" & Lno & ":" & Lin
End Function
