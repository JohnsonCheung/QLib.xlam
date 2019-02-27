VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const Ns$ = "MdyPj"
Public Md As CodeModule
Private A() As ActLin
Friend Function Init(Md As CodeModule, ActLin() As ActLin) As ActMd
Set Me.Md = Md
A = ActLin
Set Init = Me
End Function
Function ActLinAy() As ActLin()
ActLinAy = A
End Function
Function ToStr$()
Stop '
'ToStr = FmtQQ("{ActMd:{IsIns:?, Lno:?, Lin:""?""}}", IsIns, Lno, Lin)
End Function
