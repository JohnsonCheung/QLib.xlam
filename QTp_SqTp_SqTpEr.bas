Attribute VB_Name = "QTp_SqTp_SqTpEr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MTp_SqyRslt_41_ErzSqLy."
Private Const Asm$ = "QTp"
Function ErzSqTp(SqTp$) As String()

End Function

Function ErzSqLy(SqLy$()) As LyRslt

End Function
Private Function MsgAp_Lin_TyEr(DroLLin()) As String()


End Function

Private Function MsgMustBeIntoLin(DroLLin())

End Function

Private Function MsgMustBeSelorSelDis$(DroLLin())

End Function

Private Function MsgMustNotHasSpcInTbl_NmOfIntoLin(DroLLin())

End Function
Function BlkIx%(B As Blk)
BlkIx = B.DroBlk(3)
End Function
Private Function ErzExcessBlk(B As Blks, BlkTy$) As String()
Dim M As Blk: M = BlkzTy(B, BlkTy)
If IsBlkEmp(M) Then Exit Function
PushI ErzExcessBlk, FmtQQ("Excess [?] block, they are ignored", BlkTy)
PushI ErzExcessBlk, ErzAftBlk(B, M)
End Function
ErzBlk
Private Function MsgzLeftOvrAftEvl(A() As SwLin, Sw As Sw) As String()
'If Si(A) = 0 Then Exit Function
Dim I
PushI MsgzLeftOvrAftEvl, "Following lines cannot be further evaluated:"
'For Each I In A
'    PushI MsgzLeftOvrAftEvl, vbTab & CvSwLin(I).Lin
'Next
'PushIAy MsgzLeftOvrAftEvl, FmtDicTit(Sw, "Following is the [Sw] after evaluated:")
End Function

Private Function MsgzNoNm$(A As SwLin)
'MsgzNoNm = SwLinMsg(A, "No name")
End Function

Private Function MsgzOpStrEr$(A As SwLin)
MsgzOpStrEr = SwLinMsg(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function

Private Function MsgzPfx$(A As SwLin)
MsgzPfx = SwLinMsg(A, "First Char must be @")
End Function
Private Function SwLinMsg$(A As SwLin, ParamArray Ap())

End Function
Private Function MsgzTermCntAndOr$(A As SwLin)
MsgzTermCntAndOr = SwLinMsg(A, "When 2nd-Term (Operator) is [AND OR], at least 1 term")
End Function

Private Function MsgzTermCntEqNe$(A As SwLin)
MsgzTermCntEqNe = SwLinMsg(A, "When 2nd-Term (Operator) is [EQ NE], only 2 terms are allowed")
End Function

Private Function MsgzTermMustBegWithQuestOrAt$(TermAy, A As SwLin)
MsgzTermMustBegWithQuestOrAt = SwLinMsg(A, "Terms[" & JnSpc(TermAy) & "] must begin with either [?] or [@?]")
End Function

Private Function MsgzTermNotInPm$(TermAy, A As SwLin)
MsgzTermNotInPm = SwLinMsg(A, "Terms[" & JnSpc(TermAy) & "] begin with [@?] must be found in Pm")
End Function

Private Function MsgzTermNotInSw$(TermAy, A As SwLin, SwNm As Dictionary)
MsgzTermNotInSw = SwLinMsg(A, "Terms[" & JnSpc(TermAy) & "] begin with [?] must be found in Switch")
End Function

Private Function ErzDupNm(A() As SwLin, O() As SwLin) As String()
Dim Ny$(), Nm$
Dim J%, M As SwLin, Er() As SwLin
'For J = 0 To UB(A)
    'Set M = A(J)
    If HasEle(Ny, M.Nm) Then
        'PushObj Er, M
    Else
        'PushObj O, M
        PushI Ny, M.Nm
    End If
'Next
'ErzDupNm = MsgzDupNm(Er)
End Function

Private Function ErzFld(A() As SwLin, O() As SwLin) As String()
Exit Function
Dim M As SwLin, IsEr As Boolean, J%, I, A1() As SwLin, A2() As SwLin
IsEr = True
A1 = A
While IsEr
    J = J + 1: If J > 1000 Then Stop
    IsEr = False
'    For Each I In Itr(A1)
'        Set M = ErzFldLin(CvSwLin(I), SwNmDic(A), OEr)
'        If IsNothing(M) Then
            IsEr = True
'        Else
            'PushObj A2, M
'        End If
    'Next
    A1 = A2
Wend
'ErzFld = A2
End Function

Private Function ErzFldLin(A As SwLin, SwNm As Dictionary, Pm As Dictionary, OEr$()) As SwLin
'Each Term in A.TermAy must be found either in Sw or Pm
Dim O0$(), O1$(), O2$(), I
For Each I In Itr(A.TermAy)
    Select Case True
    Case HasPfx(I, "?"):  If Not SwNm.Exists(I) Then Push O0, I
    Case HasPfx(I, "@?"): If Not Pm.Exists(I) Then Push O1, I
    Case Else:                  Push O2, I
    End Select
Next
PushIAy OEr, MsgzTermNotInSw(O0, A, SwNm)
PushIAy OEr, MsgzTermNotInPm(O1, A)
PushIAy OEr, MsgzTermMustBegWithQuestOrAt(O2, A)
'If HasElezInSomAyzOfAp(O0, O1, O2) Then Set ErzFldLin = A
End Function

Private Function ErzLin1$(IO As SwLin)
Dim SwLinMsg$
With IO
    If .Nm = "" Then ErzLin1 = MsgzNoNm(IO): Exit Function
'    Select Case .OpStr
'    Case "OR", "AND": If Si(.TermAy) = 0 Then ErzLin1 = MsgzTermCntAndOr(IO): Exit Function
'    Case "EQ", "NE":  If Si(.TermAy) <> 2 Then ErzLin1 = MsgzTermCntEqNe(IO): Exit Function
'    Case Else:        ErzLin1 = MsgzOpStrEr(IO): Exit Function
'    End Select
End With
End Function

Function ErzPm(Pm As Dictionary) As String()
'Dim O As Dictionary: Set O = Dic(PmLy)
'Dim B As Boolean, K, V
'For Each K In O.Keys
'    Select Case True
'    Case HasPfx(K, ">?"):
'        V = O(K)
'        Select Case V
'        Case "0": B = False
'        Case "1": B = True
'        Case Else: Thw CSub, "If K='>?xxx', V should be 0 or 1", "K V PmLy", K, V, PmLy
'        End Select
'        O(K) = B
'    Case HasPfx(K, ">")
'    Case Else: Thw CSub, "Pm line should beg with (>? | >)", "K V PmLy", K, O(K), PmLy
'    End Select
'Next
'Set SqPm = O
End Function

