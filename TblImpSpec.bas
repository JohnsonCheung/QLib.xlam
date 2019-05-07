VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TblImpSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "TblImpSpec."
Public Tbl$, LnkColStr$, WhBexpr$
Friend Property Get Init(Tbl$, LnkColStr$, Optional WhBexpr$)
Me.Tbl = Tbl
Me.LnkColStr = LnkColStr
Me.WhBexpr = WhBexpr
Set Init = Me
End Property
