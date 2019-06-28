Attribute VB_Name = "QIde_B_Cnst"
Option Explicit
Option Compare Text
Type CnstBrk
    IsPrv As Boolean
    Cnstn As String
    TyChr As String
    Str As String
End Type
Type LLin
    Lno As Integer
    Lin As String
End Type
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
'Ret : All cnst lin as :ContLin in @Src
Dim L, Ix&: For Each L In Itr(Src)
    If IsLinCnst(L) Then PushI CnstLy, ContLin(Src, Ix)
    Ix = Ix + 1
Next
End Function

Private Function FoCnst() As String()
FoCnst = SyzSS("Mdy IsPrv Cnstn TyChr Str")
End Function

Private Function DroCnst(Lin, Mdn) As String()
'Fm Lin : Assume the lin is a :ContLin
'Ret    : :Sy<Mdn IsPrv Nm TyChr Str>: or &EmpSy if @CnstLin is not a cnst lin
With BrkCnst(Lin)
    If .Cnstn <> "" Then
    DroCnst = Sy(Mdn, .IsPrv, .Cnstn, .TyChr, .Str)
    End If
End With
End Function
Sub Z_BrkCnst()
Dim Lin, Act As CnstBrk, Ept As CnstBrk
GoSub T0
Exit Sub
T0:
    Lin = "Private Const AA$ = ""sdf"""
    Ept = CnstBrk(True, "AA", "$", """sdf""")
    GoTo Tst
Tst:
    Act = BrkCnst(Lin)
    If Not IsEqCnstBrk(Act, Ept) Then Stop
    Return
End Sub

Function IsEqCnstBrk(A As CnstBrk, B As CnstBrk) As Boolean
With A
    If .IsPrv <> B.IsPrv Then Exit Function
    If .Cnstn <> B.Cnstn Then Exit Function
    If .Str <> B.Str Then Exit Function
    If .TyChr <> B.TyChr Then Exit Function
End With
IsEqCnstBrk = True
End Function
Function CnstBrk(IsPrv As Boolean, Cnstn$, TyChr$, Str$) As CnstBrk
With CnstBrk
    .IsPrv = IsPrv
    .Cnstn = Cnstn
    .TyChr = TyChr
    .Str = Str
End With
End Function
Function BrkCnst(Lin) As CnstBrk
'Fm Lin : Assume the lin is a :ContLin
'Ret : :Ty:CnstBrk
Dim O As CnstBrk
Dim L$: L = Lin
O.IsPrv = ShfShtMdy(L) = "Prv"
If Not ShfCnst(L) Then Exit Function
O.Cnstn = ShfNm(L): If O.Cnstn = "" Then Exit Function
O.TyChr = ShfTyChr(L)
If Not ShfPfx(L, " = ") Then Exit Function
O.Str = L
BrkCnst = O
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
Function CnstLnozMN(M As CodeModule, Cnstn$) As LLin
Dim J&, L$
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    If HasPfx(L, "Const CMod$") Then
        CnstLnozMN = LLin(J - 1, L)
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
Function HitCnstn(SrcLin, Cnstn$) As Boolean
HitCnstn = CnstnzL(SrcLin) = Cnstn
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



Function SzStrCnstLin$(Lin)
Dim A As CnstBrk: A = BrkCnst(Lin)
If A.TyChr <> "$" Then Exit Function
SzStrCnstLin = SzQvbs(A.Str)
Stop
End Function
Function SzStrCnstn$(Lin, Cnstn$)
Dim A As CnstBrk: A = BrkCnst(Lin)
If A.Cnstn <> Cnstn Then Exit Function
If A.TyChr <> "$" Then Exit Function
SzStrCnstn = SzQvbs(A.Str)
End Function
Function DoCnstP() As Drs
DoCnstP = DoCnstzP(CPj)
End Function

Private Function DoCnstzP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, DoCnstzM(C.CodeModule))
    Debug.Print C.Name
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

Private Sub Z_SzStrCnstLin()
Dim O$()
Dim L: For Each L In SrczP(CPj)
    PushNB O, SzStrCnstLin(L)
    If Si(O) > 0 Then Stop
Next
BrwAy O
End Sub


