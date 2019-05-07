VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MdPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "MdPos."
Public Md As CodeModule, Pos As LinPos
Friend Function Init(Md As CodeModule, Pos As LinPos) As MdPos
Set Me.Md = Md
Set Me.Pos = Pos
Set Init = Me
End Function


