VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofBdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private F$()
Public LR As eLR
Enum eLR
    eLeft
    eRight
End Enum
Friend Function Init(A As eLR, Fny$()) As LofAli
F = Fny
LR = A
Set Init = Me
End Function
Property Get Fny() As String()
Fny = F
End Property

