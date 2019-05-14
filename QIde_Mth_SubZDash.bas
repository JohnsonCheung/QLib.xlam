Attribute VB_Name = "QIde_Mth_SubZDash"
Option Explicit
Private Const CMod$ = "MIde_Mth_SubZDash."
Private Const Asm$ = "QIde"
Function MthLinyzSubZDashMd(A As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinyzMd(A))
    If IsSubZDashMthLin(MthLin) Then PushI MthLinyzSubZDashMd, MthLin
Next
End Function
Function MthnyzSubZDashMd(A As CodeModule) As String()
MthnyzSubZDashMd = MthnyzMthLiny(MthLinyzSubZDashMd(A))
End Function


Function IsSubZDashMthLin(MthLin) As Boolean

End Function
