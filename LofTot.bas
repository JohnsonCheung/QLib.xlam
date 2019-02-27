VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofTot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Calc As XlCalculation
Private F$()
Friend Function Init(A As XlCalculation, Fny$()) As LofAli
F = Fny
Calc = A
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property
