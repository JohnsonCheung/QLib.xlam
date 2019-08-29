Attribute VB_Name = "QIde_Chk_ChkIde"
Option Explicit
Option Compare Text
Function ChkMdn(P As VBProject, Mdn) As Boolean
If HasMdnzP(P, Mdn) Then Exit Function
MsgBox FmtQQ("Mdn not found: ?|In Pj: ?", Mdn, P.Name), vbCritical
ChkMdn = True
End Function


'
