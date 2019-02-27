VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidFb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Fbn$, T$, Fb$, Fset As Aset, Bexpr$
Friend Function Init(Fbn, T, Fset As Aset, Bexpr$, Fb) As LidFb
With Me
    .Fbn = Fbn
    .Fb = Fb
    .T = T
    Set .Fset = Fset
    .Bexpr = Bexpr
End With
Set Init = Me
End Function
