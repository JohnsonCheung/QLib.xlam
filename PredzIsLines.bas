VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredzIsLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IPred
Private Const CMod$ = "PredzIsLines."
Private Function IPred_Pred(V As Variant) As Boolean
IPred_Pred = IsLines(V)
End Function
