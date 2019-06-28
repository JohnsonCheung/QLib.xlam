Attribute VB_Name = "QIde_B_MthOp"
Option Compare Text
Option Explicit
Private Const Asm$ = ""
Private Const CMod$ = "MIde_Mth_Op."
Function IsLinAsg(L) As Boolean
'Note: [Dr(NCol) = DicId(K)] is determined as Asg-lin
Dim A$: A = LTrim(L)
ShfPfxSpc A, "Set"
If ShfDotNm(A) = "" Then Exit Function
If FstChr(A) = "(" Then
    A = AftBkt(A)
End If
IsLinAsg = T1(A) = "="
End Function

Function RplMth(M As CodeModule, Mthn, NewL$) As Boolean
'Ret : True if Replaced.  Will dlt if NewL=''
Dim Lno&: Lno = MthLnozMM(M, Mthn)
If Not HasMthzM(M, Mthn) Then
    RplMth = True
    If NewL <> "" Then
        M.InsertLines M.CountOfLines + 1, NewL '<=== added at end
    End If
    Exit Function
End If
Dim OldL$: OldL = MthLzM(M, Mthn)
If OldL = NewL Then Exit Function '<== no chg
RplMth = True
RmvMth M, Mthn '<== rmv
M.InsertLines Lno, NewL '<== and ins
End Function

Sub RmvMthzMNn(M As CodeModule, Mthnn)
Dim I
For Each I In TermAy(Mthnn)
    RmvMthzMN M, I
Next
End Sub

Sub RmvMth(M As CodeModule, Mthn)
RmvMthzMN M, Mthn
End Sub

Sub RmvMthzMN(M As CodeModule, Mthn)
With MthLnoC2(M, Mthn)
    If .S2 > 0 Then M.DeleteLines .S2, .C2
    If .S1 > 0 Then M.DeleteLines .S1, .C1
End With
End Sub

Sub CpyMthAs(M As CodeModule, Mthn, AsMthn)
If HasMthzM(M, AsMthn) Then Inf CSub, "AsMth exist.", "Mdn FmMth AsMth", Mdn(M), Mthn, AsMthn: Exit Sub
End Sub
Sub BrwMd(Md As CodeModule)
If Md.CountOfLines = 0 Then BrwStr "No lines in Md[" & Mdn(Md) & "]": Exit Sub
Brw Src(Md), "Md" & Mdn(Md)
End Sub
Private Sub Z_RmvMthzMN()
Dim Md As CodeModule
Const Mthn$ = "ZZRmv1"
Dim Bef$(), Aft$()
Crt:
        Set Md = TmpMod
        ApdLines Md, LineszVbl("|'sdklfsdf||'dsklfj|Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
Tst:
        Bef = Src(Md)
        RmvMthzMN Md, Mthn
        Aft = Src(Md)

Insp:   Insp CSub, "RmvMth Test", "Bef RmvMth Aft", Bef, Mthn, Aft
Rmv:    RmvMd Md
End Sub


Sub MovMthzNM(Mthn, ToMdn)
MovMthzMNM CMd, Mthn, Md(ToMdn)
End Sub

Sub MovMthzMNM(Md As CodeModule, Mthn, ToMd As CodeModule)
CpyMth Mthn, Md, ToMd
RmvMthzMN Md, Mthn
End Sub

Function CdzEmpFun$(FunNm)
CdzEmpFun = FmtQQ("Function ?()|End Function", FunNm)
End Function

Function CdzEmpSub$(Subn)
CdzEmpSub = FmtQQ("Sub ?()|End Sub", Subn)
End Function
Sub AddSub(Subn)
ApdLines CMd, CdzEmpSub(Subn)
JmpMth Subn
End Sub

Sub AddFun(FunNm)
ApdLines CMd, CdzEmpFun(FunNm)
JmpMth FunNm
End Sub

Function CpyMd(FmM As CodeModule, ToM As CodeModule) As Boolean
'@FmM & @ToM must exist @@
CpyMd = RplMd(ToM, SrcL(FmM))
End Function

Function CpyMth(Mthn, FmM As CodeModule, ToM As CodeModule) As Boolean
Dim NewL$
'NewL
    NewL = MthLzM(FmM, Mthn)
'Rpl
    CpyMth = RplMth(ToM, Mthn, NewL)
'
'Const CSub$ = CMod & "CpyMdMthToMd"
'Dim Nav(): ReDim Nav(2)
'GoSub BldNav: ThwIf_ObjNE Md, ToMd, CSub, "Fm & To md cannot be same", Nav
'If Not HasMthzM(Md, Mthn) Then GoSub BldNav: ThwNav CSub, "Fm Mth not Exist", Nav
'If HasMthzM(ToMd, Mthn) Then GoSub BldNav: ThwNav CSub, "To Mth exist", Nav
'ToMd.AddFromString MthLzM(Md, Mthn)
'RmvMth Md, Mthn
'If Not IsSilent Then Inf CSub, FmtQQ("Mth[?] in Md[?] is copied ToMd[?]", Mthn, Mdn(Md), Mdn(ToMd))
'Exit Sub
'BldNav:
'    Nav(0) = "FmMd Mth ToMd"
'    Nav(1) = Mdn(Md)
'    Nav(2) = Mthn
'    Nav(3) = Mdn(ToMd)
'    Return
End Function

Function CpyMthAsVer(M As CodeModule, Mthn, Ver%) As Boolean
'Ret True if copied
Dim VerMthn$, NewL$, L$, OldL$
If Not HasMthzM(M, Mthn) Then Inf CSub, "No from-mthn", "Md Mthn", Mdn(M), Mthn: Exit Function
VerMthn = Mthn & "_Ver" & Ver
'NewL
    L = MthLzM(M, Mthn)
    NewL = Replace(L, "Sub " & Mthn, "Sub " & VerMthn, Count:=1)
'Rpl
    CpyMthAsVer = RplMth(M, VerMthn, NewL)

End Function

