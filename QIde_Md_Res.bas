Attribute VB_Name = "QIde_Md_Res"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Md_Res."
Private Const Asm$ = "QIde"
Function ResLyMd(M As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Dim Z$
    Z = MthLineszM(M, ResPfx & ResNm)
    If Si(Z) = 0 Then
        Thw CSub, "Mthn not found", "Mthn Md ResNm ResPfx", ResPfx & ResNm, Mdn(M), ResNm, ResPfx
    End If
    Z = AyeFstEle(Z)
    Z = AyeLasEle(Z)
'ResLyMd = RmvFstChrzAy(Z)
End Function

Function ReStrMd$(M As CodeModule, ResNm$)
ReStrMd = JnCrLf(ResLyMd(M, ResNm))
End Function

