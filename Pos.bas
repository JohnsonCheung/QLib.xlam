VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "Pos."
Public Cno1&, Cno2&
Friend Function Init(Cno1, Cno2) As Pos
If Cno1 > 0 Then Me.Cno1 = Cno1
If Cno2 > 0 Then Me.Cno2 = Cno2
Set Init = Me
End Function
