Attribute VB_Name = "MIde_CmpPrp"
Option Explicit
Function PrpNyzCmp(A As VBComponent) As String()
PrpNyzCmp = Itn(A.Properties)
End Function
