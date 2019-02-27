VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TpPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum eTpPosTy
    ePosRCC = 1
    ePosRR = 2
    ePosR = 3
End Enum
Public Ty As eTpPosTy
Public R1 As Integer
Public R2 As Integer
Public C1 As Integer
Public C2 As Integer
Property Get Lin$()
Dim O$
Select Case Ty
Case ePosRCC
    O = FmtQQ("RCC(? ? ?) ", R1, C1, C2)
Case ePosRR
    O = FmtQQ("RR(? ?) ", R1, R2)
Case ePosR
    O = FmtQQ("R(?)", R1)
Case Else
    'Thw CSub TpPos_FmtStr", "Invalid {TpPos}", A.Ty
End Select
End Property

