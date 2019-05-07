VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "LinPos."
Public Lno&, Pos As Pos
Friend Function Init(Lno, Pos As Pos) As LinPos
Me.Lno = Lno
Set Me.Pos = Pos
Set Init = Me
End Function

