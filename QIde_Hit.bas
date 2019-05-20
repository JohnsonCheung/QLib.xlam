Attribute VB_Name = "QIde_Hit"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Hit."
Private Const Asm$ = "QIde"

Function WhMthKd(S) As String()

End Function
Function WhMthMdyPm(A As Dictionary) As String()
PushNonBlank WhMthMdyPm, A.SwNm("Pub")
PushNonBlank WhMthMdyPm, A.SwNm("Prv")
PushNonBlank WhMthMdyPm, A.SwNm("Frd")
End Function

Function WhMthMdy(WhStr$) As String()
WhMthMdy = WhMthMdyPm(Lpm(WhStr, C_WhMthSpec))
End Function

Function HitCmp(A As VBComponent, B As WhMd) As Boolean
HitCmp = True
If HitAy(A.Type, B.CmpTy) Then
    If HitNm(A.Name, B.WhNm) Then
        Exit Function
    End If
End If
HitCmp = False
End Function
