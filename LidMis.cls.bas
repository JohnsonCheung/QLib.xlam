VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidMis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Ffn As Aset
Private A_Ty() As LidMisTy
Private A_Col() As LidMisCol
Private A_Tbl() As LidMisTbl
Friend Function Init(Ffn As Aset, Tbl() As LidMisTbl, Col() As LidMisCol, Ty() As LidMisTy) As LidMis
Set Me.Ffn = Ffn
A_Ty = Ty
A_Tbl = Tbl
A_Col = Col
Set Init = Me
End Function
Property Get Ty() As LidMisTy()
Ty = A_Ty
End Property
Property Get Tbl() As LidMisTbl()
Tbl = A_Tbl
End Property
Property Get Col() As LidMisCol()
Col = A_Col
End Property
