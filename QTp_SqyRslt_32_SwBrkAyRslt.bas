Attribute VB_Name = "QTp_SqyRslt_32_SwBrkAyRslt"
Option Explicit
Private Const CMod$ = "MTp_SqyRslt_32_SwBrkAyRslt."
Private Const Asm$ = "QTp"
Type SwBrkAyRslt: Er() As String: SwBrkAy() As SwBrk: End Type
Private Function SwBrkAyRsltzEr(SwBrkAy() As SwBrk, Er$()) As SwBrkAyRslt
SwBrkAyRsltzEr.Er = Er
SwBrkAyRsltzEr.SwBrkAy = SwBrkAy
End Function
Function SwBrkAyRslt(A() As SwBrk, Pm As Dictionary) As SwBrkAyRslt
Dim Ok1() As SwBrk
Dim Ok2() As SwBrk
Dim Ok3() As SwBrk
Dim Ok4() As SwBrk
Dim Er$(): Er = AddAyAp(ErzLin(A, Ok1), ErzDupNm(Ok1, Ok2), ErzFld(Ok2, Ok3), ErzLeftOvr(Ok3, Ok4))
SwBrkAyRslt = SwBrkAyRsltzEr(Ok4, Er)
End Function

Private Function Msgz$(A As SwBrk, B$)
Msgz = A.Lin & " --- " & B
End Function
Private Function MsgzDupNm(A() As SwBrk) As String()
Dim I
For Each I In Itr(A)
    PushI MsgzDupNm, Msgz(CvSwBrk(I), "Dup name")
Next
End Function

Private Function MsgzLeftOvrAftEvl(A() As SwBrk, Sw As Dictionary) As String()
If Si(A) = 0 Then Exit Function
Dim I
PushI MsgzLeftOvrAftEvl, "Following lines cannot be further evaluated:"
For Each I In A
    PushI MsgzLeftOvrAftEvl, vbTab & CvSwBrk(I).Lin
Next
PushIAy MsgzLeftOvrAftEvl, FmtDicTit(Sw, "Following is the [Sw] after evaluated:")
End Function

Private Function MsgzNoNm$(A As SwBrk)
MsgzNoNm = Msgz(A, "No name")
End Function

Private Function MsgzOpStrEr$(A As SwBrk)
MsgzOpStrEr = Msgz(A, "2nd Term [Op] is invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Function

Private Function MsgzPfx$(A As SwBrk)
MsgzPfx = Msgz(A, "First Char must be @")
End Function

Private Function MsgzTermCntAndOr$(A As SwBrk)
MsgzTermCntAndOr = Msgz(A, "When 2nd-Term (Operator) is [AND OR], at least 1 term")
End Function

Private Function MsgzTermCntEqNe$(A As SwBrk)
MsgzTermCntEqNe = Msgz(A, "When 2nd-Term (Operator) is [EQ NE], only 2 terms are allowed")
End Function

Private Function MsgzTermMustBegWithQuestOrAt$(TermSy$(), A As SwBrk)
MsgzTermMustBegWithQuestOrAt = Msgz(A, "Terms[" & JnSpc(TermSy) & "] must begin with either [?] or [@?]")
End Function

Private Function MsgzTermNotInPm$(TermSy$(), A As SwBrk)
MsgzTermNotInPm = Msgz(A, "Terms[" & JnSpc(TermSy) & "] begin with [@?] must be found in Pm")
End Function

Private Function MsgzTermNotInSw$(TermSy$(), A As SwBrk, SwNm As Dictionary)
MsgzTermNotInSw = Msgz(A, "Terms[" & JnSpc(TermSy) & "] begin with [?] must be found in Switch")
End Function

Private Function ErzDupNm(A() As SwBrk, O() As SwBrk) As String()
Dim Ny$(), Nm$
Dim J%, M As SwBrk, Er() As SwBrk
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

Private Function ErzFld(A() As SwBrk, O() As SwBrk) As String()
Exit Function
Dim M As SwBrk, IsEr As Boolean, J%, I, A1() As SwBrk, A2() As SwBrk
IsEr = True
A1 = A
While IsEr
    J = J + 1: If J > 1000 Then Stop
    IsEr = False
    For Each I In Itr(A1)
'        Set M = ErzFldLin(CvSwBrk(I), SwNmDic(A), OEr)
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

Private Function ErzFldLin(A As SwBrk, SwNm As Dictionary, Pm As Dictionary, OEr$()) As SwBrk
'Each Term in A.TermSy must be found either in Sw or Pm
Dim O0$(), O1$(), O2$(), I
For Each I In Itr(A.TermSy)
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

Private Function ErzLeftOvr(A() As SwBrk, O() As SwBrk) As String()
End Function

Private Function ErzLin1$(IO As SwBrk)
Dim Msgz$
With IO
    If .Nm = "" Then ErzLin1 = MsgzNoNm(IO): Exit Function
    Select Case .OpStr
    Case "OR", "AND": If Si(.TermSy) = 0 Then ErzLin1 = MsgzTermCntAndOr(IO): Exit Function
    Case "EQ", "NE":  If Si(.TermSy) <> 2 Then ErzLin1 = MsgzTermCntEqNe(IO): Exit Function
    Case Else:        ErzLin1 = MsgzOpStrEr(IO): Exit Function
    End Select
End With
End Function

Private Function ErzLin(A() As SwBrk, O() As SwBrk) As String()
Dim I
For Each I In Itr(A)
    PushNonNothing ErzLin, ErzLin1(CvSwBrk(I))
Next
End Function

Private Function ErzPfx(A() As Lnx, OEr$()) As Lnx()
Dim J%
For J = 0 To UB(A)
    If FstChr(A(J).Lin) <> "?" Then
        PushI OEr, MsgzPfx(A(J))
    Else
        PushObj ErzPfx, A(J)
    End If
Next
End Function

