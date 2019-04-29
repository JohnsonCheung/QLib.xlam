VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredzIsNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IPred
Private Function IPred_Pred(V As Variant) As Boolean
If Not IsStr(V) Then Exit Function
IPred_Pred = IsNm(CStr(V))
End Function
