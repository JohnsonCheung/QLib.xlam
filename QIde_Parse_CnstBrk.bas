Attribute VB_Name = "QIde_Parse_CnstBrk"
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Const."
Public Type StrCnst
    Nm As String
    V As String
End Type
Public Type CnstBrk ' It comes from a ConstLin
    IsPrv As Boolean
    Nm As String
    TyChr As String
    AsTy As String 'Eg Date, Boolean, Only either TyChr or AsTy will have value
    Val As String
    Rmk As String
End Type
Public Type SomCnstBrk ' It comes a Lin
    Som As Boolean
    Itm As CnstBrk
End Type
Public Type StrCnsts:    N As Long: Ay() As StrCnst:    End Type
Public Type CnstBrks:    N As Long: Ay() As CnstBrk:    End Type
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
    If .Str <> "" Then
        PushStrCnst O, StrCnst(.Nm, .Str)
    End If
    End With
Next
End Function
Function StrCnst(Nm$, Str$) As StrCnst
StrCnst.Nm = Nm
StrCnst.Str = Str
End Function
Sub PushStrCnst(O As StrCnsts, M As StrCnst)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function CnstBrks(Ly$()) As CnstBrks
Dim L$, I, O As CnstBrks
For Each I In Itr(ContLyzLy(Ly))
    L = I
    CnstBrksPushOpt O, SomCnstBrk(L)
Next
CnstBrks = O
End Function
Sub CnstBrksPushOpt(O As CnstBrks, M As SomCnstBrk)
If Not M.Som Then Exit Sub
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M.V
O.N = O.N + 1
End Sub
Function SomCnstBrk(Lin$) As SomCnstBrk
Dim L$: L = Lin
Dim IsPrv As Boolean: IsPrv = ShfMdy(L) = "Private"
Select Case True
Case Not ShfPfx(L, "Const") Or (ShfNm(L) = "")
Case Not ShfPfx(L, "$")
Case Not ShfPfx(L, " = """)
Case Else
    Dim P%:
    P = InStr(L, """"): If P = 0 Then Thw CSub, "Should  have 2 dbl-quote", "Lin", Lin
    ValzCnstLin = Left(L, P - 1)
    Exit Function
End Select
CnstBrk(
Exit Function


Dim L$, N$: L = Lin
Dim O As SomCnstBrk
O.V.IsPrv = ShfMdy(L) = "Private"
If Not ShfPfx(L, "Const ") Then Exit Function
O.V.Nm = ShfNm(L)
O.V.TyChr = ShfTyChr(L)
If ShfTermX(L, "As") Then
    O.V.Ty = ShfT1(L)
    O.V.NonStrVal = L
    Exit Function
End If
If Not ShfTermX(L, "=") Then Thw CSub, "Lin is invalid const line: no [ = ] after name", "Lin", Lin
If ShfPfx(L, """") Then
    Dim P&: P = InStr(L, """"): If P = 0 Then Thw CSub, "Something wrong in Lin, which is supposed to be string const lin.  There is no snd [""]", "Lin", Lin
    O.V.Str = Left(L, P - 1)
Else
    O.V.NonStrVal = L
End If
O.Som = True
SomCnstBrk = O
End Function
Function StrCnstszBrks(A As CnstBrks) As StrCnsts
Dim J%, V$
For J = 0 To A.N - 1
    With A.Ay(J)
        If .TyChr = "$" Or .AsTy = "String" Then
            V = TakVbStr(.Val)
            PushStrCnst StrCnstszBrks, StrCnst(.Nm, V)
        End If
    End With
Next
End Function
Function StrCnsts(Ly$()) As StrCnsts
StrCnsts = StrCnstzCnstBrks(CnstBrks(Ly))
End Function


Function CnstNm$(Lin$)
Dim L$: L = Lin
ShfMdy L
If ShfPfx(L, "Const") Then CnstNm = Nm(L)
End Function


Function ShfConst(OLin$) As Boolean
ShfConst = ShfPfx(OLin, "Const")
End Function

Private Sub Z_HasCnstNm()
Debug.Assert HasCnstNm(CurMd, "CMod")
End Sub
Function HasCnstNm(A As CodeModule, CnstNm$) As Boolean
Dim J%
For J = 1 To A.CountOfDeclarationLines
    If HitCnstNm(A.Lines(J, 1), CnstNm) Then HasCnstNm = True: Exit Function
Next
End Function

Function CnstNmzSrcLin$(SrcLin$)
Dim L$: L = RmvMdy(SrcLin)
If ShfConst(L) Then CnstNmzSrcLin = Nm(LTrim(L))
End Function


