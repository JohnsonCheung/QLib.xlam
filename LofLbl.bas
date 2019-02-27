VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofLbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fld$, Lbl$
Friend Function Init(Fld$, Lbl$) As LofFml
Me.Lbl = Lbl
Me.Fld = Fld
Set Init = Me
End Function


