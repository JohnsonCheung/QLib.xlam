VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Nm As String, IsOpt As Boolean, IsPmAy As Boolean, IsAy As Boolean, TyChr$, AsTy$, DftVal As String
Friend Function Init(ArgStr$) As Arg
Set Init = Me
End Function
Property Get ToStr$()
Dim O$()
ToStr = JnSpc(O)
End Property

Property Get ShtStr$()
End Property
