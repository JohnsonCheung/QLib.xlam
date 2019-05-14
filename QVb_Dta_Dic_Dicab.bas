Attribute VB_Name = "QVb_Dta_Dic_Dicab"
Public Type Dicab
    A As Dictionary
    B As Dictionary
End Type
Function Dicab(A As Dictionary, B As Dictionary) As Dicab
ThwIf_Nothing A, CSub
ThwIf_Nothing B, CSub
End Function
