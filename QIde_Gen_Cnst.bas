Attribute VB_Name = "QIde_Gen_Cnst"
Option Explicit
Option Compare Text
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Const."
Type StrCnst
    Nm As String
    S As String
End Type
Type CnstBrk ' It comes from a ConstLin
    IsPrv As Boolean
    Nm As String
    TyChr As String
    AsTy As String 'Eg Date, Boolean, Only either TyChr or AsTy will have value
    Val As String
    Rmk As String
End Type
Type SomCnstBrk ' It comes a Lin
    Som As Boolean
    Itm As CnstBrk
End Type
Type StrCnsts:    N As Long: Ay() As StrCnst:    End Type
Type CnstBrks:    N As Long: Ay() As CnstBrk:    End Type
Public Const DoczCnstVal$ = "It is CnstBrk.Val"
Function StrValzCnstBrk$(A As CnstBrk)
If IsStrCnst(A) Then StrValzCnstBrk = StrValzCnstVal(A.Val)
End Function
Private Function IsStrCnst(A As CnstBrk) As Boolean

End Function
Private Function StrValzCnstVal(CnstVal$)

End Function
Function StrCnstszCnstBrks(A As CnstBrks) As StrCnsts
Dim O As StrCnsts, J%
For J = 0 To A.N - 1
    With A.Ay(J)
    If .TyChr = "$" Or .AsTy = "String" Then
        PushItrCnst O, StrCnst(.Nm, .Val)
    End If
    End With
Next
End Function
Function StrCnst(Nm, S) As StrCnst
StrCnst.Nm = Nm
StrCnst.S = S
End Function
Sub PushItrCnst(O As StrCnsts, M As StrCnst)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function CnstLy(Src$()) As String()

End Function

Function CnstBrks(Src$()) As CnstBrks
Dim L
For Each L In Itr(CnstLy(Src))
    PushCnstBrk CnstBrks, CnstBrk(L)
Next
End Function

Sub PushCnstBrk(O As CnstBrks, M As CnstBrk)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function SomCnstBrk(Lin) As SomCnstBrk
Dim L$: L = Lin
Dim IsPrv As Boolean: IsPrv = ShfMdy(L) = "Private"
Select Case True
Case Not ShfPfx(L, "Const") Or (ShfNm(L) = "")
Case Not ShfPfx(L, "$")
Case Not ShfPfx(L, " = """)
Case Else
    Dim P%:
    P = InStr(L, """"): If P = 0 Then Thw CSub, "Should  have 2 dbl-quote", "Lin", Lin
    'ValzCnstLin = Left(L, P - 1)
    Exit Function
End Select
Exit Function
'
'
'Dim N$: L = Lin
'Dim O As SomCnstBrk
'O.V.IsPrv = ShfMdy(L) = "Private"
'If Not ShfPfx(L, "Const ") Then Exit Function
'O.V.Nm = ShfNm(L)
'O.V.TyChr = ShfTyChr(L)
'If ShfTerm(L, "As") Then
'    O.V.Ty = ShfT1(L)
'    O.V.NonStrVal = L
'    Exit Function
'End If
'If Not ShfTerm(L, "=") Then Thw CSub, "Lin is invalid const line: no [ = ] after name", "Lin", Lin
'If ShfPfx(L, """") Then
'    Dim P&: P = InStr(L, """"): If P = 0 Then Thw CSub, "Something wrong in Lin, which is supposed to be string const lin.  There is no snd [""]", "Lin", Lin
'    O.V.Str = Left(L, P - 1)
'Else
'    O.V.NonStrVal = L
'End If
'O.Som = True
'SomCnstBrk = O
End Function
Function StrCnstszBrks(A As CnstBrks) As StrCnsts
Dim J%, V$
For J = 0 To A.N - 1
    With A.Ay(J)
        If .TyChr = "$" Or .AsTy = "String" Then
            'V = TakVbStr(.Val)
            PushItrCnst StrCnstszBrks, StrCnst(.Nm, V)
        End If
    End With
Next
End Function
Function StrCnsts(Ly$()) As StrCnsts
'StrCnsts = StrCnstzCnstBrks(CnstBrks(Ly))
End Function

Function Cnstn$(Lin)
Dim L$: L = Lin
ShfMdy L
If ShfPfx(L, "Const") Then Cnstn = TakNm(L)
End Function

Function CnstLnozMN(M As CodeModule, Cnstn$) As Lnx
Dim J&, L$
For J = 1 To M.CountOfDeclarationLines
    L = M.Lines(J, 1)
    If HasPfx(L, "Const CMod$") Then
        CnstLnozMN = Lnx(L, J - 1)
        Exit Function
    End If
Next
End Function

Function ShfTermCnst(OLin$) As Boolean
ShfTermCnst = ShfTerm(OLin, "Const")
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

Function CnstnzSrcLin(SrcLin)
Dim L$: L = RmvMdy(SrcLin)
If ShfTermCnst(L) Then CnstnzSrcLin = Nm(LTrim(L))
End Function

Function CnstBrk(CnstLin) As CnstBrk

End Function
Function DrzStrCnst(Lin) As Variant()
Dim L$: L = RmvMdy(Lin)
If Not ShfConst(L) Then Exit Function
Dim N$: N = ShfNm(L): If N = "" Then Exit Function
If Not ShfPfx(L, "$") Then Exit Function
If Not ShfPfx(L, " = """) Then Exit Function
Dim P%: P = InStr(L, """"): If P = 0 Then Stop
DrzStrCnst = Array(N, Left(L, P - 1))
End Function
Function StrValzCnstn$(Lin, Cnstn$)
Dim L$: L = RmvMdy(Lin)
If Not ShfConst(L) Then Exit Function
If ShfNm(L) <> Cnstn$ Then Exit Function
If Not ShfPfx(L, "$") Then Exit Function
If Not ShfPfx(L, " = """) Then Stop
Dim P%: P = InStr(L, """"): If P = 0 Then Stop
StrValzCnstn = Left(L, P - 1)
End Function
Function DStrCnstP() As Drs
DStrCnstP = DStrCnst(SrczP(CPj))
End Function
Function DStrCnst(Src$()) As Drs
Dim ODry(), L
For Each L In Itr(Src)
    PushISomSi ODry, DrzStrCnst(L)
Next
DStrCnst = DrszFF("Cnstn StrVal", ODry)
End Function

Function StrValzCnstLy$(Ly$(), Cnstn$)
Dim L
For Each L In Itr(Ly)
    Dim O$: O = StrValzCnstn(L, Cnstn)
    If O <> "" Then StrValzCnstLy = O: Exit Function
Next
End Function
Function StrValzCnstLin(Lin)
StrValzCnstLin = StrValzCnstBrk(CnstBrk(Lin))
End Function

Function CMCnstLy(CMSrc$()) As String()
Dim L
For Each L In Itr(CMSrc)
PushI CMCnstLy, CMCnstLin(L)
Next
End Function
Function CMCnstLin$(CMSrcLin)
Dim N, T1$, L$, O$
L = CMSrcLin
T1 = ShfT1(L)
O = "Private Const C_" & T1 & "$ = """ & L
For Each N In NyzMacro(CMSrcLin)
    O = Replace(O, QuoteBigBkt(N), "?")
Next
CMCnstLin = O & """"
End Function

Function CMFunLinesAy(CMSrc$()) As String()
Dim L
For Each L In Itr(CMSrc)
PushI CMFunLinesAy, CMFunLines(L)
Next
End Function
Function CMFunLines$(CMSrcLin)
If InStr(CMSrcLin, "{") = 0 Then Exit Function
Dim O$(), Nm$, Pm$, PmOnlyNm$, Ny$(), NyOnlyNm$()
Nm = T1(CMSrcLin)
Ny = AywDist(NyzMacro(CMSrcLin))
Pm = JnComma(Ny)
'NyOnlyNm = TakNm zAy(Ny)
PmOnlyNm = JnComma(NyOnlyNm)
PushI O, FmtMacro("Private Function M_{Nm}$({Pm})", Nm, Pm)
PushI O, FmtMacro("M_{Nm} = FmtQQ(C_{Nm}, {PmNmOnly})", Nm, PmOnlyNm)
PushI O, "End Function"
End Function

