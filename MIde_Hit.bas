Attribute VB_Name = "MIde_Hit"
Option Explicit

Function WhMthKd(S) As String()

End Function
Function WhMthMdyPm(A As LinPm) As String()
PushNonBlankStr WhMthMdyPm, A.SwNm("Pub")
PushNonBlankStr WhMthMdyPm, A.SwNm("Prv")
PushNonBlankStr WhMthMdyPm, A.SwNm("Frd")
End Function

Function WhMthMdy(WhStr$) As String()
WhMthMdy = WhMthMdyPm(LinPm(WhStr))
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
