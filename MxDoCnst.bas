Attribute VB_Name = "MxDoCnst"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoCnst."
Public Const FFoCnst$ = "Mdn Mdy Cnstn TyChr CnstLin"

Function FoCnst() As String()
FoCnst = SyzSS(FFoCnst)
End Function

Function DoCnstP() As Drs
DoCnstP = DoCnstzP(CPj)
End Function

Function DoCnstzP(P As VBProject) As Drs
Dim O()
Dim C As VBComponent, M As CodeModule: For Each C In P.VBComponents
    Set M = C.CodeModule
    PushI O, DyoCnst(Src(M), M.Name)
Next
DoCnstzP = Drs(FoCnst, O)
End Function

Sub LisCnst()

End Sub

Function DoCnst(Src$(), Mdn$) As Drs
DoCnst = Drs(FoCnst, DyoCnst(Src, Mdn))
End Function

Private Function DyoCnst(Src$(), Mdn$) As Variant()
Dim L: For Each L In Itr(Src)
    PushSomSi DyoCnst, DroCnst(L, Mdn)
Next
End Function

Private Sub Z_CnstLy()
Brw CnstLy(SrczP(CPj))
End Sub

Function CnstLy(Src$()) As String()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsLinCnst(Src(Ix)) Then PushI CnstLy, ContLin(Src, Ix)
Next
End Function

Function DroCnst(Lin, Optional Mdn$) As Variant()
'Ret    : :Dro|EmpAv if @Lin is not a cnst-cont-lin
Dim L$: L = Lin
Dim IsPrv As Boolean: IsPrv = ShfShtMdy(L) = "Prv"
If Not ShfCnst(L) Then Exit Function
Dim Cnstn$: Cnstn = ShfNm(L): If Cnstn = "" Then Exit Function
Dim TyChr$: TyChr = ShfTyChr(L)
If Not ShfPfx(L, " = ") Then Exit Function
DroCnst = Array(Mdn, IsPrv, Cnstn, TyChr, L)
End Function

