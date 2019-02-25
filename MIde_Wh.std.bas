Attribute VB_Name = "MIde_Wh"
Option Explicit

Function WhMthzStr(WhStr$) As WhMth
Dim ShtMdy$(), ShtKd$(), Nm As WhNm
Dim A As LinPm: Set A = LinPm(WhStr)
With A
    PushNonBlankStr ShtMdy, .SwNm("Pub")
    PushNonBlankStr ShtMdy, .SwNm("Prv")
    PushNonBlankStr ShtMdy, .SwNm("Frd")
    PushNonBlankStr ShtKd, .SwNm("Sub")
    PushNonBlankStr ShtKd, .SwNm("Fun")
    PushNonBlankStr ShtKd, .SwNm("Prp")
    Set Nm = WhNmzStr(WhStr)
End With
Set WhMthzStr = WhMth(ShtMdy, ShtKd, Nm)
End Function

Function WhMdMth(Optional Md As WhMd, Optional Mth As WhMth) As WhMdMth
Set WhMdMth = New WhMdMth
With WhMdMth
    Set .Md = Md
    Set .Mth = Mth
End With
End Function

Function WhMdMth_WhMd(A As WhMdMth) As WhMd
If Not IsNothing(A) Then Set WhMdMth_WhMd = A.Md
End Function

Function WhMdMth_WhMth(A As WhMdMth) As WhMth
If Not IsNothing(A) Then Set WhMdMth_WhMth = A.Mth
End Function

Function WhMth(ShtMdy$(), ShtKd$(), Nm As WhNm) As WhMth
Dim O As New WhMth
Set WhMth = O.Init(ShtMdy, ShtKd, Nm)
End Function

Function WhPjMth(Optional Pj As WhNm, Optional MdMth As WhMdMth) As WhPjMth
Set WhPjMth = New WhPjMth
With WhPjMth
    Set .Pj = Pj
    Set .MdMth = MdMth
End With
End Function
Function WhNm(Patn$, LikAy$(), ExlLikAy$()) As WhNm
Dim O As New WhNm
Set WhNm = O.Init(Patn, LikAy, ExlLikAy)
End Function

Function WhMd(CmpTy() As vbext_ComponentType, Nm As WhNm) As WhMd
Set WhMd = New WhMd
WhMd.Init CmpTy, Nm
End Function
Function WhMdzStr(WhStr$) As WhMd
With LinPm(WhStr)
    Dim CmpTy() As vbext_ComponentType
    If .HasSw("Cls") Then PushI CmpTy, vbext_ct_ClassModule
    If .HasSw("Mod") Then PushI CmpTy, vbext_ct_StdModule
    Set WhMdzStr = WhMd(CmpTy, .WhNm)
End With
End Function


