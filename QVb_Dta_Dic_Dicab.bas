Attribute VB_Name = "QVb_Dta_Dic_Dicab"
Option Explicit
Option Compare Text
Type Dicab
    A As Dictionary
    B As Dictionary
End Type
Function Dicab(A As Dictionary, B As Dictionary) As Dicab
ThwIf_Nothing A, "DicA", CSub
ThwIf_Nothing B, "DicB", CSub
End Function
