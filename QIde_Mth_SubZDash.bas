Attribute VB_Name = "QIde_Mth_SubZDash"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_SubZDash."
Private Const Asm$ = "QIde"
Function MthLinAyzSubZDashMd(M As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinAyzM(M))
    If IsSubZDashMthLin(MthLin) Then PushI MthLinAyzSubZDashMd, MthLin
Next
End Function
Function MthnyzSubZDashMd(M As CodeModule) As String()
MthnyzSubZDashMd = MthnyzMthLinAy(MthLinAyzSubZDashMd(M))
End Function


Function IsSubZDashMthLin(MthLin) As Boolean

End Function
