Attribute VB_Name = "MxCnst"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCnst."
Public Const FFoCnst$ = "Mdn IsPrv Cnstn, TyChr PrimStr"
Enum XoCnst
    XiMdn
    XiIsPrv
    XiCnstn
    XiTyChr
    XiCnstv
End Enum

Function CnstLyP() As String()
CnstLyP = CnstLyzP(CPj)
End Function

Function CnstLyzP(P As VBProject) As String()
CnstLyzP = CnstLy(SrczP(P))
End Function

Function Cnstn$(Lin)
Dim L$: L = RmvMdy(Lin)
If ShfPfxSpc(L, "Const") Then
    Cnstn = TakNm(L)
End If
End Function

Function CnstLno%(M As CodeModule, Cnstn$, Optional IsPrvOnly As Boolean)
CnstLno = CnstIx(Src(M), Cnstn, IsPrvOnly) + 1
End Function

Function CnstLLin(M As CodeModule, Cnstn$) As LLin
Dim J&: For J = 1 To M.CountOfDeclarationLines
    Dim L$: L = M.Lines(J, 1)
    If CnstnzL(L) = Cnstn Then
        L = ContLinzM(M, J)
        CnstLLin = LLin(J, L)
        Exit Function
    End If
Next
End Function

Private Sub Z_HasCnstn()
Debug.Assert HasCnstn(CMd, "CMod")
End Sub

Function HasCnstn(M As CodeModule, Cnstn$) As Boolean
HasCnstn = CnstLno(M, Cnstn) = 0
End Function

Function HitCnstn(Lin, Cnstn$) As Boolean
HitCnstn = CnstnzL(Lin) = Cnstn
End Function

Function CnstnzL$(L)
CnstnzL = Cnstn(L)
End Function

Function ShfTermCnst(OLin$) As Boolean
ShfTermCnst = ShfTerm(OLin, "Const")
End Function

Function ShfCnst(OLin$) As Boolean
ShfCnst = ShfT1(OLin) = "Const"
End Function

Function CnstStr$(Lin)
CnstStr = Ele(DroCnst(Lin, ""), XiMdn)
End Function

Function CnstStrzN$(Lin, Cnstn$)
Dim A(): A = DroCnst(Lin, ""): If Si(A) = 0 Then Exit Function
If A(XiCnstn) <> Cnstn Then Exit Function
If A(XiTyChr) <> "$" Then Exit Function
CnstStrzN = UnQteVb(A(XiCnstv))
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

Function CnstIx&(Src$(), Cnstn, Optional IsPrvOnly As Boolean)
Dim L, O&
For Each L In Itr(Src)
    If CnstnzL(L) = Cnstn Then
        Select Case True
        Case IsPrvOnly And HasPfx(L, "Public "): CnstIx = -1
        Case Else:                              CnstIx = O
        End Select
        Exit Function
    End If
    O = O + 1
Next
CnstIx = -1
End Function

Function CnstLinAy(Src$()) As String()
Dim Ix&, L: For Each L In Itr(Src)
    If IsLinCnst(L) Then PushI CnstLinAy, ContLin(Src, Ix)
    Ix = Ix + 1
Next
End Function

Function CnstLinAyP() As String()
CnstLinAyP = CnstLinAy(SrczP(CPj))
End Function

Function CnstFeizM(M As CodeModule, Cnstn$) As Fei
CnstFeizM = CnstFei(DclzM(M), Cnstn)
End Function

Function CnstFei(Dcl$(), Cnstn$) As Fei
Dim J&: For J = 0 To UB(Dcl)
    If HitCnstn(Dcl(J), Cnstn) Then
        CnstFei = ContFei(Dcl, J)
        Exit Function
    End If
Next
End Function