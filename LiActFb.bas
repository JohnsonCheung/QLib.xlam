VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiActFb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Fb As String
Public Fbn As String
Public T As String
Public Fset As Aset
Friend Function Init(Fb$, Fbn$, T$, Fset As Aset) As LiActFb
With Me
    .Fb = Fb
    .Fbn = Fbn
    Set .Fset = Fset
    .T = T
End With
Set Init = Me
End Function
