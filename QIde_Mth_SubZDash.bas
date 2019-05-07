Attribute VB_Name = "QIde_Mth_SubZDash"
Option Explicit
Private Const CMod$ = "MIde_Mth_SubZDash."
Private Const Asm$ = "QIde"
Function MthLinSyzSubZDashMd(A As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinSyzMd(A))
    If IsSubZDashMthLin(MthLin) Then PushI MthLinSyzSubZDashMd, MthLin
Next
End Function
Function MthNyzSubZDashMd(A As CodeModule) As String()
MthNyzSubZDashMd = MthNyzMthLinSy(MthLinSyzSubZDashMd(A))
End Function


Function IsSubZDashMthLin(MthLin) As Boolean

End Function
