VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofTit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Tit$, Fld$
Friend Function Init(Fld$, Tit$) As LofFml
Me.Tit = Tit
Me.Fld = Fld
Set Init = Me
End Function



