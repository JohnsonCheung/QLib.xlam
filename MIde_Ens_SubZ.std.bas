Attribute VB_Name = "MIde_Ens_SubZ"
Option Explicit

Private Function SubZEptzNy$(MthNySubZDash$()) ' Sub Z() bodylines
Dim O$()
PushI O, "Private Sub Z()"
PushIAy O, AySrt(MthNySubZDash)
PushI O, "End Sub"
SubZEptzNy = JnCrLf(O)
End Function

Function SubZEptzMd$(A As CodeModule)
'SubZ is [Mth-`Sub Z()`-Lines], each line is calling a Z_XX, where Z_XX is a testing function
SubZEptzMd = SubZEptzNy(MthNyzSubZDashMd(A))
End Function

