VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActPj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A() As ActMd
Friend Function Init(ActMd() As ActMd) As ActPj
A = ActMd
End Function
Function ActMdAy() As ActMd()
ActMdAy = A
End Function
