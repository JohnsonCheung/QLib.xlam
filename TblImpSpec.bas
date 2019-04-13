VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TblImpSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Tbl$, LnkColStr$, WhBExpr$
Friend Property Get Init(Tbl$, LnkColStr$, Optional WhBExpr$)
Me.Tbl = Tbl
Me.LnkColStr = LnkColStr
Me.WhBExpr = WhBExpr
Set Init = Me
End Property
