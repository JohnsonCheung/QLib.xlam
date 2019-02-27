VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofAli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private F$()
Public Ali As XlHAlign
Friend Function Init(A As XlHAlign, Fny$()) As LofAli
F = Fny
Ali = A
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property
