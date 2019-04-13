Attribute VB_Name = "MIde_Md_Res"
Option Explicit
Function ResLyMd(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Dim Z$
    Z = MthLinesByMdMth(A, ResPfx & ResNm)
    If Si(Z) = 0 Then
        Thw CSub, "MthNm not found", "MthNm Md ResNm ResPfx", ResPfx & ResNm, MdNm(A), ResNm, ResPfx
    End If
    Z = AyeFstEle(Z)
    Z = AyeLasEle(Z)
ResLyMd = AyRmvFstChr(Z)
End Function

Function ReStrMd$(A As CodeModule, ResNm$)
ReStrMd = JnCrLf(ResLyMd(A, ResNm))
End Function

