Attribute VB_Name = "MIde_Mth_SubZDash"
Option Explicit
Function MthLinAyzSubZDashMd(A As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinAyzMd(A))
    If IsSubZDashMthLin(MthLin) Then PushI MthLinAyzSubZDashMd, MthLin
Next
End Function
Function MthNyzSubZDashMd(A As CodeModule) As String()
MthNyzSubZDashMd = MthNyzMthLinAy(MthLinAyzSubZDashMd(A))
End Function


Function IsSubZDashMthLin(MthLin) As Boolean

End Function
