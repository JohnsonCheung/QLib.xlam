Attribute VB_Name = "QIde_Ens_EnsAsm"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const Ns$ = "QIde.Qualify"
Private Const CMod$ = "BEnsAsm."

Function AsmnzMdn$(Mdn$)
If FstChr(Mdn) = "Q" Then
    If HasSubStr(Mdn, "_") Then
        AsmnzMdn = Bef(Mdn, "_")
    End If
End If
End Function

Function DrszMapAy(Ay, MapFunNN$, Optional FF$) As Drs
Dim Dry(), V: For Each V In Ay
    Dim Dr(): Dr = Array(V)
    Dim F: For Each F In Itr(SyzSS(MapFunNN))
        PushI Dr, Run(F, V)
    Next
    PushI Dry, Dr
Next
Dim A$: A = DftStr(FF, "V " & MapFunNN)
DrszMapAy = DrszFF(A, Dry)
End Function

Sub EnsAsmP()
EnsAsmzP CPj
End Sub

Function EnsAsmzM(M As CodeModule) As Boolean
If IsEmpMd(M) Then Exit Function
If CmpTyzM(M) = vbext_ct_Document Then Exit Function

Const T$ = "Private Const CMod$ = ""?"""
Dim N$, L&, C$, Mdn$
Mdn = MdnzM(M)
   C = "CMod"
  N = MdnzM(M)
  L = LnozDclCnst(M, C)
If L = 0 Then L = LnozFstCd(M)
If EnsLin(M, L, FmtQQ(T, C, N & ".")) Then EnsAsmzM = True

     C = "Asm"
     N = AsmnzMdn(Mdn)
         If EnsAsmzM__1(M, C, N, T) Then EnsAsmzM = True

     C = "Ns"
     N = NsnzMdn(Mdn)
         If EnsAsmzM__1(M, C, N, T) Then EnsAsmzM = True
End Function
Private Function EnsAsmzM__1(M As CodeModule, C$, N$, T$) As Boolean
Dim L&, Lin$, NoNm As Boolean, HasLin As Boolean
     L = LnozDclCnst(M, C)
   Lin = FmtQQ(T, C, N)
  NoNm = N = ""
HasLin = L <> 0
If Not (NoNm And HasLin) Then _
    EnsAsmzM__1 = EnsLin(M, L, Lin)

End Function
Sub EnsAsmzP(P As VBProject)
Dim C As VBComponent, Mdyd%, Skpd%
For Each C In P.VBComponents
    If EnsAsmzM(C.CodeModule) Then
        Mdyd = Mdyd + 1
    Else
        Skpd = Skpd + 1
    End If
Next
Inf CSub, "Done", "Pj Mdyd Skpd Tot", P.Name, Mdyd, Skpd, Mdyd + Skpd
End Sub

Sub EnsCnstzMth(M As CodeModule, Mthn$, Cnstn$, NewL$)

End Sub

Function EnsLin(M As CodeModule, L&, NewL$) As Boolean
If L = 0 Then Exit Function
If M.Lines(L, 1) = NewL Then Exit Function
If NewL = "" Then
    M.DeleteLines L, 1
Else
    M.ReplaceLine L, NewL
End If
EnsLin = True
End Function

Function HasAsmn(Mdn) As Boolean
If FstChr(Mdn) <> "M" Then Exit Function
If Not IsAscUCas(Asc(SndChr(Mdn))) Then Exit Function
HasAsmn = True
End Function

Function IxzCnst&(Src$(), Cnstn$)
Dim O&, S
For Each S In Itr(Src)
    If CnstnzL(S) = Cnstn Then
        IxzCnst = O
    End If
    O = O + 1
Next
IxzCnst = -1
End Function

Function LnozDclCnst%(M As CodeModule, Cnstn$)
Dim O%, L$
Dim C$: C = "Const " & Cnstn
For O = 1 To M.CountOfDeclarationLines
    L = RmvMdy(M.Lines(O, 1))
    If ShfPfx(L, "Const ") Then
        If TakNm(L) = Cnstn Then LnozDclCnst = O: Exit Function
    End If
Next
End Function

Function LnozFstCd&(M As CodeModule)
Stop

End Function

Function NsnzMdn$(Mdn$)
If FstChr(Mdn) = "Q" Then
    Dim A$: A = BefOrAll(Mdn, "__")
    Dim P1%: P1 = InStr(A, "_")
    If P1 = 0 Then Exit Function
    Dim P2%: P2 = InStrRev(A, "_")
    If P1 = P2 Then Exit Function
    NsnzMdn = Mid(A, P1 + 1, P2 - P1 - 1)
End If
End Function

Private Sub Z_LnozDclConst()
Dim Md As CodeModule, Cnstn$
GoSub T0
Exit Sub
T0:
    Set Md = CMd
    Cnstn = "A$"
    Ept = 14&
    GoTo Tst
Tst:
    Act = LnozDclCnst(Md, Cnstn)
    C
    Return
End Sub

Sub ZZ_AsmnzMdn()
BrwDrs DrszMapAy(Itn(CPj.VBComponents), "AsmnzMdn NsnzMdn")
End Sub
