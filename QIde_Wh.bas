Attribute VB_Name = "QIde_Wh"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Wh."
Private Const Asm$ = "QIde"
Public Const C_WhMthSpec$ = ""
Function WhMthzPfx(WhMthnPfx$, Optional InclPrv As Boolean) As WhMth

End Function
Function WhMthzSfx(WhMthnSfx$, Optional InclPrv As Boolean) As WhMth

End Function

Function WhMthzStr(WhStr$) As WhMth
Dim ShtMdy$(), ShtKd$()
Const C$ = ""
Dim A As Dictionary: Set A = Lpm(WhStr, C)
With A
    PushNonBlank ShtMdy, .SwNm("Pub")
    PushNonBlank ShtMdy, .SwNm("Prv")
    PushNonBlank ShtMdy, .SwNm("Frd")
    PushNonBlank ShtKd, .SwNm("Sub")
    PushNonBlank ShtKd, .SwNm("Fun")
    PushNonBlank ShtKd, .SwNm("Prp")
End With
WhMthzStr = WhMth(ShtMdy, ShtKd, WhNmzS(WhStr))
End Function

Function WhMdMth(WhMd As WhMd, WhMth As WhMth) As WhMdMth
With WhMdMth
    .WhMd = WhMd
    .WhMth = WhMth
End With
End Function

Function WhMdzWhMdMth(A As WhMdMth) As WhMd
'If Not IsNothing(A) Then Set WhMdzWhMdMth = A.Md
End Function

Function WhMthzWhMdMth(A As WhMdMth) As WhMth
'If Not IsNothing(A) Then Set WhMthzWhMdMth = A.Mth
End Function

Function WhMth(ShtMdy$(), ShtKd$(), Nm As WhNm) As WhMth
With WhMth
    
End With
End Function

Function WhPjMth(Pj As WhNm, Md As WhMdMth) As WhPjMth
With WhPjMth
    .WhPjNm = Pj
    .WhMdMth = Md
End With
End Function

Function WhNm(Patn$, LikAy$(), ExlLikAy$()) As WhNm
With WhNm
    .ExlLikAy = ExlLikAy
    .LikAy = LikAy
    If Patn <> "" Then Set .Re = RegExp(Patn)
End With
End Function

Function WhMd(CmpTy() As vbext_ComponentType, Nm As WhNm) As WhMd
'Set WhMd = New WhMd
'WhMd.Init CmpTy, Nm
End Function
Function WhMdzStr(WhStr$) As WhMd
With Lpm(WhStr, C_WhMthSpec)
    Dim CmpTy() As vbext_ComponentType
    If .HasSw("Cls") Then PushI CmpTy, vbext_ct_ClassModule
    If .HasSw("Mod") Then PushI CmpTy, vbext_ct_StdModule
'    WhMdzStr = WhMd(CmpTy, .WhNm)
End With
End Function


