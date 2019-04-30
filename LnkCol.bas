VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LnkCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Nm$, Ty As Dao.DataTypeEnum, Extnm$
Friend Property Get Init(Nm, Ty As Dao.DataTypeEnum, Extnm$)
Me.Nm = Nm
Me.Ty = Ty
Me.Extnm = Extnm
Set Init = Me
End Property
