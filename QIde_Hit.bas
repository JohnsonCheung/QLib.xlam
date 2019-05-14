Attribute VB_Name = "QIde_Hit"
Option Explicit
Private Const CMod$ = "MIde_Hit."
Private Const Asm$ = "QIde"

Function WhMthKd(S) As String()

End Function
Function WhMthMdyPm(A As Lpm) As String()
PushNonBlank WhMthMdyPm, A.SwNm("Pub")
PushNonBlank WhMthMdyPm, A.SwNm("Prv")
PushNonBlank WhMthMdyPm, A.SwNm("Frd")
End Function

Function WhMthMdy(WhStr$) As String()
WhMthMdy = WhMthMdyPm(Lpm(WhStr, C_WhMthSpec))
End Function

Function HitCmp(A As VBComponent, B As WhMd) As Boolean
HitCmp = True
If IsNothing(B) Then Exit Function
If HitAy(A.Type, B.CmpTy) Then
    If HitNm(A.Name, B.Nm) Then
        Exit Function
    End If
End If
HitCmp = False
End Function
