VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredzIsSy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IPred
Private Const CMod$ = "PredzIsSy."
Private Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = IsSy(V)
End Function
