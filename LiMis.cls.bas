VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiMis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public MisFfn As Aset
Private A_Col() As LiMisCol
Private A_Ty() As LiMisTy
Private A_Tbl() As LiMisTbl
Friend Function Init(MisFfn As Aset, MisTbl() As LiMisTbl, MisCol() As LiMisCol, MisTy() As LiMisTy) As LiMis
Set Me.MisFfn = MisFfn
A_Tbl = MisTbl
A_Col = MisCol
A_Ty = MisTy
Set Init = Me
End Function
Property Get MisTy() As LiMisTy()
MisTy = A_Ty
End Property
Property Get MisTbl() As LiMisTbl()
MisTbl = A_Tbl
End Property
Property Get MisCol() As LiMisCol()
MisCol = A_Col
End Property

