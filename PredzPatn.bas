VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredzPatn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IPred
Private A As New RegExp
Friend Sub Init(Patn$)
A.Pattern = Patn
End Sub
Private Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = A.Test(V)
End Function
