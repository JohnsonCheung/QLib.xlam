Attribute VB_Name = "QIde_Ens_SubZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZ."
Private Const Asm$ = "QIde"

Private Function SubZEptzNy$(MthnySubZDash$()) ' Sub Z() bodylines
Dim O$()
PushI O, "Private Sub ZZ()"
PushIAy O, SrtAy(MthnySubZDash)
PushI O, "End Sub"
SubZEptzNy = JnCrLf(O)
End Function

Function SubZEptzMd$(M As CodeModule)
'SubZ is [Mth-`Sub Z()`-Lines], each line is calling a Z_XX, where Z_XX is a testing function
SubZEptzMd = SubZEptzNy(MthnyzSubZDashMd(M))
End Function

