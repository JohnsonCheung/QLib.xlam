VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredHasPatn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IPred
Private Re As RegExp
Friend Sub Init(Patn$)
Set Re = RegExp(Patn)
End Sub
Function Pred(V) As Boolean
Pred = IPred_Pred(V)
End Function
Private Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = Re.Test(V)
End Function

