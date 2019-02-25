VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidFil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public FilNm$, Ffn$
Friend Function Init(FilNm$, Ffn$) As LidFil
With Me
    .FilNm = FilNm
    .Ffn = Ffn
End With
Set Init = Me
End Function
