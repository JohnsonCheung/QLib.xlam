VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiFil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FilNm$, Ffn$
Friend Function Init(FilNm, Ffn) As LiFil
With Me
    .FilNm = FilNm
    .Ffn = Ffn
End With
Set Init = Me
End Function
