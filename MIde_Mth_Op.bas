Attribute VB_Name = "MIde_Mth_Op"
Option Explicit
Const CMod$ = "MIde_Mth_Op."

Function TmpMod() As CodeModule
Set TmpMod = AddCmp(TmpModNm, vbext_ct_StdModule).CodeModule
End Function

Sub ClrTmpMod()
Dim C As VBComponent
For Each C In CurPj.VBComponents
    If HasPfx(C.Name, "TmpMod") Then C.Delete
Next
End Sub

Property Get TmpModNm$()
TmpModNm = "TmpMod" & NowNm
End Property

Function NowNm$()
Static N&
NowNm = Format(Now, "HHMMSS") & N
N = N + 1
End Function

Sub RmvMth(A As CodeModule, MthNmNN)
Dim MthNm
For Each MthNm In TermAy(MthNmNN)
    RmvMthzNm A, MthNm
Next
End Sub

Sub RmvMthzNm(A As CodeModule, MthNm, Optional WiTopRmk As Boolean)
RmvFTIxAy A, MthFTIxAyzMth(A, MthNm, WiTopRmk)
End Sub

Sub RmvMdMth(Md As CodeModule, MthNm)
Const CSub$ = CMod & "RmvMdMth"
Dim X() As FTIx: X = MthFTIxAyzMth(Md, MthNm)
Inf CSub, "Remove method", "Md Mth FTIx-WiTopRmk", MdNm(Md), Md, LyzFTIxAy(X)
RmvMdLineszFtIxAy Md, X
End Sub

Private Sub Z_RmvMdMth()
Const N$ = "ZZModule"
Dim Md As CodeModule, MthNm$
GoSub Crt
'RmvMdEndBlankLin Md
Tst:
    RmvMdMth Md, MthNm
    If Md.CountOfLines <> 0 Then Stop
    Return
Crt:
    Set Md = TmpMod
    RmvMd Md
    ApdLines Md, LineszVbl("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Private Sub Z()
End Sub
Sub CpyMdMthToMd(Md As CodeModule, MthNm, ToMd As CodeModule, Optional IsSilent As Boolean)
Const CSub$ = CMod & "CpyMdMthToMd"
Dim Nav(): ReDim Nav(2)
GoSub BldNav: ThwEqObjNav Md, ToMd, CSub, "Fm & To md cannot be same", Nav
If Not HasMthMd(Md, MthNm) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
If HasMthMd(ToMd, MthNm) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
ToMd.AddFromString MthLinesByMdMth(Md, MthNm)
RmvMdMth Md, MthNm
If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", MthNm, MdNm(Md), MdNm(ToMd))
Exit Sub
BldNav:
    Nav(0) = "FmMd Mth ToMd"
    Nav(1) = MdNm(Md)
    Nav(2) = MthNm
    Nav(3) = MdNm(ToMd)
    Return
End Sub

Sub MovMthTo(MthNm, ToMdNm$)
MovMdMthTo CurMd, MthNm, Md(ToMdNm)
End Sub

Sub MovMdMthTo(Md As CodeModule, MthNm, ToMd As CodeModule)
CpyMdMthToMd Md, MthNm, ToMd
RmvMdMth Md, MthNm
End Sub

