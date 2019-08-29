Attribute VB_Name = "QIde_B_Cnst"
Option Explicit
Option Compare Text
Type LLin
    Lno As Integer
    Lin As String
End Type
Const FFoCnst$ = "IsPrv Cnstn TyChr Str Mdn"
Enum IxoCnst
    XoIsPrv = 0
    XoCnstn = 1
    XoTyChr = 2
    XoStr = 3
    XoMdn = 4
End Enum

Private Function DoCnst(Src$(), Mdn) As Drs
DoCnst = Drs(FoCnst, DyoCnst(Src, Mdn))
End Function

Private Function DyoCnst(Src$(), Mdn) As Variant()
Dim L: For Each L In Itr(Src)
    PushSomSi DyoCnst, DroCnst(L, Mdn)
Next
End Function

Private Sub Z_CnstLy()
Brw CnstLy(SrczP(CPj))
End Sub

Private Function CnstLy(Src$()) As String()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsLinCnst(Src(Ix)) Then PushI CnstLy, ContLin(Src, Ix)
Next
End Function

Private Function FoCnst() As String()
FoCnst = SyzSS(FFoCnst)
End Function

Private Function DroCnst(Lin, Mdn) As Variant()
'Fm Lin : Assume the lin is a :ContLin
'Ret    : :Sy<Mdn IsPrv Nm TyChr Str>: or &EmpSy if @CnstLin is not a cnst lin
'Fm Lin : Assume the lin is a :ContLin
'Ret : :Ty:CnstBrk
Dim L$: L = Lin
Dim IsPrv As Boolean: IsPrv = ShfShtMdy(L) = "Prv"
If Not ShfCnst(L) Then Exit Function
Dim Cnstn$: Cnstn = ShfNm(L): If Cnstn = "" Then Exit Function
Dim TyChr$: TyChr = ShfTyChr(L)
If Not ShfPfx(L, " = ") Then Exit Function
DroCnst = Array(IsPrv, Cnstn, TyChr, L, Mdn)
End Function

Function CnstLyP() As String()
CnstLyP = CnstLyzP(CPj)
End Function

Function CnstLyzP(P As VBProject) As String()
CnstLyzP = CnstLy(SrczP(P))
End Function

Function Cnstn$(Lin)
Dim L$: L = Lin
ShfMdy L
If ShfPfx(L, "Const") Then Cnstn = TakNm(L)
End Function

Function LLin(Lno&, Lin$) As LLin
LLin.Lno = Lno
LLin.Lin = Lin
End Function

Function CnstLLin(M As CodeModule, Cnstn$) As LLin
Dim J&, L$
Dim Dcl$(): Dcl = DclLyzM(M)
Dim Ix%: For Ix = 0 To UB(Dcl)
    L = Dcl(Ix)
    If HasPfx(L, "Const CMod$") Then
        CnstLLin = LLin(J - 1, ContLin(Dcl, Ix))
        Exit Function
    End If
Next
End Function

Private Sub Z_HasCnstn()
Debug.Assert HasCnstn(CMd, "CMod")
End Sub
Function HasCnstn(M As CodeModule, Cnstn$) As Boolean
Dim J%
For J = 1 To M.CountOfDeclarationLines
    If HitCnstn(M.Lines(J, 1), Cnstn) Then HasCnstn = True: Exit Function
Next
End Function

Function HitCnstn(Lin, Cnstn$) As Boolean
HitCnstn = CnstnzL(Lin) = Cnstn
End Function

Function CnstnzL$(L)
Dim Lin$: Lin = RmvMdy(L)
If ShfTermCnst(Lin) Then CnstnzL = Nm(LTrim(Lin))
End Function

Function ShfTermCnst(OLin$) As Boolean
ShfTermCnst = ShfTerm(OLin, "Const")
End Function

Function ShfCnst(OLin$) As Boolean
ShfCnst = ShfT1(OLin) = "Const"
End Function

Function CnstStr$(Lin)
CnstStr = Ele(DroCnst(Lin, ""), XoMdn)
End Function

Function CnstStrzN$(Lin, Cnstn$)
Dim A(): A = DroCnst(Lin, ""): If Si(A) = 0 Then Exit Function
If A(XoCnstn) <> Cnstn Then Exit Function
If A(XoTyChr) <> "$" Then Exit Function
CnstStrzN = UnQteVb(A(XoStr))
End Function

Function DoCnstP() As Drs
DoCnstP = DoCnstzP(CPj)
End Function

Private Function DoCnstzP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, DoCnstzM(C.CodeModule))
Next
DoCnstzP = O
End Function

Private Function DoCnstzM(M As CodeModule) As Drs
DoCnstzM = DoCnst(Src(M), Mdn(M))
End Function

Function IsLinCnstPfx(L, CnstnPfx$) As Boolean
Dim Lin$: Lin = RmvMdy(L)
If Not ShfTermCnst(Lin) Then Exit Function
IsLinCnstPfx = HasPfx(L, CnstnPfx)
End Function
Function IsLinCnst(L) As Boolean
IsLinCnst = T1(RmvMdy(L)) = "Const"
End Function

Private Sub Z_CnstStr()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    PushNB O, CnstStr(L)
    If Si(O) > 0 Then Stop
Next
BrwAy O
End Sub

Function CnstIx&(Src$(), Cnstn)
Dim L, O&
For Each L In Itr(Src)
    If CnstnzL(L) = Cnstn Then CnstIx = O: Exit Function
    O = O + 1
Next
CnstIx = -1
End Function


'
