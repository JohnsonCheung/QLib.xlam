Attribute VB_Name = "QIde_Md_Res"
Option Explicit
Private Const CMod$ = "MIde_Md_Res."
Private Const Asm$ = "QIde"
Function ResLyMd(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Dim Z$
    Z = MthLineszMN(A, ResPfx & ResNm)
    If Si(Z) = 0 Then
        Thw CSub, "Mthn not found", "Mthn Md ResNm ResPfx", ResPfx & ResNm, Mdn(A), ResNm, ResPfx
    End If
    Z = AyeFstEle(Z)
    Z = AyeLasEle(Z)
'ResLyMd = RmvFstChrzAy(Z)
End Function

Function ReStrMd$(A As CodeModule, ResNm$)
ReStrMd = JnCrLf(ResLyMd(A, ResNm))
End Function

