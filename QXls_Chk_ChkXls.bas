Attribute VB_Name = "QXls_Chk_ChkXls"
Option Explicit
Option Compare Text

Function IsEqCC(Lo As ListObject, CC$, Optional IsInf As Boolean) As Boolean
Dim FF$: FF = JnSpc(FnyzLo(Lo))
If FF = CC Then IsEqCC = True: Exit Function
MsgBox FmtQQ("Expected: ?|Actual: ?", CC, FF), vbCritical, FmtQQ("Lo[?] fields error", Lo.Name)
End Function

'
