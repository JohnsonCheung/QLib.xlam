Attribute VB_Name = "QIde_CmpPrp"
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_CmpPrp."
Function PrpNyzCmp(A As VBComponent) As String()
PrpNyzCmp = Itn(A.Properties)
End Function
