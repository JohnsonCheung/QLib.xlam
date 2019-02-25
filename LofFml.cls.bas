VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LofFml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fld$, Fml$
Friend Function Init(Fld$, Fml$) As LofFml
Me.Fml = Fml
Me.Fld = Fld
Set Init = Me
End Function

