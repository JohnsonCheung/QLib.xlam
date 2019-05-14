Attribute VB_Name = "QTp_SqTp_SqTpEr"
Option Explicit
Private Const CMod$ = "MTp_SqyRslt_41_ErzSqLy."
Private Const Asm$ = "QTp"
Function ErzSqTp(SqTp$) As String()

End Function

Function ErzSqLy(SqLy$()) As LyRslt

End Function
Private Function MsgAp_Lin_TyEr(A As Lnx) As String()


End Function

Private Function MsgMustBeIntoLin(A As Lnx)

End Function

Private Function MsgMustBeSelorSelDis$(A As Lnx)

End Function

Private Function MsgMustNotHasSpcInTbl_NmOfIntoLin(A As Lnx)

End Function


Private Function ErzExcessPmBlk(A As Blks) As String()
If CntBlk(A, "PM") > 1 Then
'    PushIAy ErzExcessPmBlk, ErzBlk(CvBlk(AyeFstEle(Blk)), "Excess Pm block, they are ignored")
End If
End Function

Private Function ErzExcessSwBlk(A As Blks) As String()
If CntBlk(A, "SW") > 1 Then
'    PushIAy ErzExcessSwBlk, ErzBlk(CvBlk(AyeFstEle(Blk)), "Excess Sw block, they are ignored")
End If
End Function
Blks
Function ErzBlk(Blk As Blk, Msg$) As String()

End Function

Private Function MsgzLeftOvrAftEvl(A() As SwLin, Sw As Sw) As String()
If Si(A) = 0 Then Exit Function
Dim I
PushI MsgzLeftOvrAftEvl, "Following lines cannot be further evaluated:"
For Each I In A
    PushI MsgzLeftOvrAftEvl, vbTab & CvSwLin(I).Lin
Next
PushIAy MsgzLeftOvrAftEvl, FmtDicTit(Sw, "Following is the [Sw] after evaluated:")
End Function

Private Function MsgzNoNm$(A As SwLin)
MsgzNoNm = Msgz(A, "No name")
End Function

Private Function MsgzOpStrEr$(A As SwLin)
MsgzOpStrEr = Msgz(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function

Private Function MsgzPfx$(A As SwLin)
MsgzPfx = Msgz(A, "First Char must be @")
End Function

Private Function MsgzTermCntAndOr$(A As SwLin)
MsgzTermCntAndOr = Msgz(A, "When 2nd-Term (Operator) is [AND OR], at least 1 term")
End Function

Private Function MsgzTermCntEqNe$(A As SwLin)
MsgzTermCntEqNe = Msgz(A, "When 2nd-Term (Operator) is [EQ NE], only 2 terms are allowed")
End Function

Private Function MsgzTermMustBegWithQuestOrAt$(TermAy, A As SwLin)
MsgzTermMustBegWithQuestOrAt = Msgz(A, "Terms[" & JnSpc(TermAy) & "] must begin with either [?] or [@?]")
End Function

Private Function MsgzTermNotInPm$(TermAy, A As SwLin)
MsgzTermNotInPm = Msgz(A, "Terms[" & JnSpc(TermAy) & "] begin with [@?] must be found in Pm")
End Function

Private Function MsgzTermNotInSw$(TermAy, A As SwLin, SwNm As Dictionary)
MsgzTermNotInSw = Msgz(A, "Terms[" & JnSpc(TermAy) & "] begin with [?] must be found in Switch")
End Function

Private Function ErzDupNm(A() As SwLin, O() As SwLin) As String()
Dim Ny$(), Nm$
Dim J%, M As SwLin, Er() As SwLin
For J = 0 To UB(A)
    Set M = A(J)
    If HasEle(Ny, M.Nm) Then
        PushObj Er, M
    Else
        PushObj O, M
        PushI Ny, M.Nm
    End If
Next
ErzDupNm = MsgzDupNm(Er)
End Function

Private Function ErzFld(A() As SwLin, O() As SwLin) As String()
Exit Function
Dim M As SwLin, IsEr As Boolean, J%, I, A1() As SwLin, A2() As SwLin
IsEr = True
A1 = A
While IsEr
    J = J + 1: If J > 1000 Then Stop
    IsEr = False
    For Each I In Itr(A1)
'        Set M = ErzFldLin(CvSwLin(I), SwNmDic(A), OEr)
        If IsNothing(M) Then
            IsEr = True
        Else
            PushObj A2, M
        End If
    Next
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
If HasElezInSomAyzOfAp(O0, O1, O2) Then Set ErzFldLin = A
End Function

Private Function ErzLin1$(IO As SwLin)
Dim Msgz$
With IO
    If .Nm = "" Then ErzLin1 = MsgzNoNm(IO): Exit Function
    Select Case .OpStr
    Case "OR", "AND": If Si(.TermAy) = 0 Then ErzLin1 = MsgzTermCntAndOr(IO): Exit Function
    Case "EQ", "NE":  If Si(.TermAy) <> 2 Then ErzLin1 = MsgzTermCntEqNe(IO): Exit Function
    Case Else:        ErzLin1 = MsgzOpStrEr(IO): Exit Function
    End Select
End With
End Function

