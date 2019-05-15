Attribute VB_Name = "QIde_Mth_Op"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Mth_Op."


Sub RmvMthzMNn(A As CodeModule, Mthnn, Optional WiTopRmk As Boolean)
Dim I
For Each I In TermAy(Mthnn)
    RmvMthzMN A, I
Next
End Sub

Sub RmvMth(A As CodeModule, Mthn)
RmvMthzMN A, Mthn
End Sub

Sub RmvMthzMN(A As CodeModule, Mthn)
DltLinzF A, MthFeiszMN(A, Mthn, WiTopRmk:=True)
End Sub

Private Sub Z_RmvMthzMN()
Const N$ = "ZZModule"
Dim Md As CodeModule, Mthn
GoSub Crt
'RmvMdEndBlankLin Md
Tst:
    RmvMthzMN Md, Mthn
    If Md.CountOfLines <> 0 Then Stop
    Return
Crt:
    Set Md = TmpMod
    RmvMd Md
    ApdLines Md, LineszVbl("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Private Sub ZZ()
End Sub
Sub CpyMth(Md As CodeModule, Mthn, ToMd As CodeModule, Optional IsSilent As Boolean)
Const CSub$ = CMod & "CpyMdMthToMd"
Dim Nav(): ReDim Nav(2)
GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
ToMd.AddFromString MthLineszMN(Md, Mthn)
RmvMth Md, Mthn
If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", Mthn, Mdn(Md), Mdn(ToMd))
Exit Sub
BldNav:
    Nav(0) = "FmMd Mth ToMd"
    Nav(1) = Mdn(Md)
    Nav(2) = Mthn
    Nav(3) = Mdn(ToMd)
    Return
End Sub

Sub MovMthzNM(Mthn, ToMdn)
MovMthzMNM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzMNM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Md, Mthn, ToMd
RmvMthzMN Md, Mthn
End Sub

Function EmpFunLines$(FunNm)
EmpFunLines = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function EmpSubLines$(Subn)
EmpSubLines = FmtQQ("Sub ?()|End Sub", Subn)
End Function
Sub AddSub(Subn)
ApdLines CMd, EmpSubLines(Subn)
JmpMth Subn
End Sub

Sub AddFun(FunNm)
ApdLines CMd, EmpFunLines(FunNm)
JmpMth FunNm
End Sub

