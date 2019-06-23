Attribute VB_Name = "QIde_Md_Res"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Res."
Private Const Asm$ = "QIde"
Function ResLyzM(M As CodeModule, Mthn$) As String()
Dim Z$
    Z = MthLzM(M, Mthn)
    If Si(Z) = 0 Then
        Thw CSub, "Mthn not found", "Mthn Md", Mthn, Mdn(M)
    End If
    Z = AeFstEle(Z)
    Z = AeLasEle(Z)
ResLyzM = RmvFstChrzAy(Z)
End Function

Function ReszM$(M As CodeModule, Mthn$)
ReszM = JnCrLf(ResLyzM(M, Mthn))
End Function

