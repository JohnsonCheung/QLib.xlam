VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredzIsPrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Private Const CMod$ = "PredzIsPrim."
Private Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = IsPrim(V)
End Function
