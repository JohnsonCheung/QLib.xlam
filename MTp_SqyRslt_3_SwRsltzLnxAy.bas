Attribute VB_Name = "MTp_SqyRslt_3_SwRsltzLnxAy"
Option Explicit
Const CMod$ = "MTp_Sq_Sw."
Private Samp As New SampSqt
Type SwRslt: Er() As String: FldSw As Dictionary: StmtSw As Dictionary: End Type

Private Function SwRsltzEr(FldSw As Dictionary, StmtSw As Dictionary, Er$()) As SwRslt
With SwRsltzEr
    .Er = Er
    Set .FldSw = FldSw
    Set .StmtSw = StmtSw
End With
End Function
Private Sub AAMain()
Z_SwRslt
End Sub

Function SwRsltzLnxAy(A() As Lnx, Pm As Dictionary) As SwRslt
Dim Inp() As SwBrk:          Inp = SwBrkAy(A)
Dim R0 As SwBrkAyRslt:        R0 = SwBrkAyRslt(Inp, Pm)
Dim Inp1() As SwBrk:        Inp1 = R0.SwBrkAy
Dim OSw As New Dictionary
    Dim NotEvl() As SwBrk
    Dim R As DicOpt
    Dim IsEvl As Boolean
    Dim I%, J%
    Dim Evling() As SwBrk:
    Evling = Inp1
    IsEvl = True
    While IsEvl
        IsEvl = False
        J = J + 1
        If J > 1000 Then Stop
        For I = 0 To UB(Evling)
            R = EvlSwBrkLin(Evling(I), R.Dic)
            If R.Som Then
                IsEvl = True
            Else
                PushObj NotEvl, Evling(I)
            End If
        Next
        If False Then
            Brw AyAddAp( _
                LblTabFmtAySepSS("*InpSwBrk", SwBrkAyFmt(Inp)), _
                LblTabFmtAySepSS("*OkSwBrk", SwBrkAyFmt(Evling)), _
                LblTabFmtAySepSS("*NotEvl", SwBrkAyFmt(NotEvl)), _
                LblTabFmtAySepSS("OSw", FmtDic(OSw)))
            Stop
        End If
        Evling = NotEvl
        Erase NotEvl
    Wend
SwRsltzLnxAy = SwRsltzSwDic(OSw, R0.Er)
End Function
Private Function SwRsltzSwDic(SwDic As Dictionary, Er$()) As SwRslt
Dim OStmt As Dictionary: Set OStmt = New Dictionary
Dim OFld As Dictionary: Set OFld = New Dictionary
Dim K
For Each K In SwDic.Keys
    If FstTwoChr(K) <> "?#" Then  ' Skip.  It is temp-Sw
        Select Case Left(K, 5)
        Case "?SEL#", "?UPD#"     ' It is StmtSw
            OStmt.Add Mid(K, 5), SwDic(K)
        Case Else                 ' It is FldSw
            OFld.Add K, SwDic(K)
        End Select
    End If
Next
End Function

Private Function EvlBoolTerm(BoolTerm, Sw As Dictionary, BoolTermPm As Dictionary) As BoolOpt
Dim O As BoolOpt
If BoolTermPm.Exists(BoolTerm) Then
    O.Bool = BoolTermPm(BoolTerm)
    O.Som = True
Else
    If Sw.Exists(BoolTerm) Then
        O.Bool = Sw(BoolTerm)
        O.Som = True
    End If
End If
EvlBoolTerm = O
End Function
Private Function EvlSwBrkLin(A As SwBrk, Sw As Dictionary) As DicOpt
'Return True and set Result if evalulated
Const CSub$ = CMod & "EvlSwBrkLin"
If Sw.Exists(A.Nm) Then Thw CSub, "[SwBrk] should not be found in [Sw]", A.Lin, FmtDic(Sw)
Dim Ay$(): Ay = A.TermAy
Dim R As BoolOpt, BoolTermPm As Dictionary
Stop
Select Case A.OpStr
Case "OR":  R = EvlTermAy(Ay, "OR", Sw, BoolTermPm)
Case "AND": R = EvlTermAy(Ay, "AND", Sw, BoolTermPm)
Case "NE":  R = EvlT1T2(Ay(0), Ay(1), "NE", Sw)
Case "EQ":  R = EvlT1T2(Ay(0), Ay(1), "EQ", Sw)
Case Else: Thw CSub, "[SwBrk] has invalid [OpStr], where [Valid OpStr]", A.Lin, A.OpStr, "OR AND NE EQ"
End Select

If R.Som Then
    Dim O As Dictionary: Set O = DicClone(Sw)
    O.Add A.Nm, R.Bool
    EvlSwBrkLin = SomDic(O)
End If
End Function


Private Function EvlT1(T1$, Sw As Dictionary, SwTermPm As Dictionary) As StrOpt
EvlT1 = EvlTerm(T1, Sw, SwTermPm)
End Function

Private Function EvlT1T2(T1$, T2$, EQ_NE$, Sw As Dictionary) As BoolOpt
'Return True and set ORslt if evaluated
Const CSub$ = CMod & "EvlT1T2"
Dim SwTermPm As Dictionary
Dim R1 As StrOpt, R2 As StrOpt
R1 = EvlT1(T1, Sw, SwTermPm)
R2 = EvlT2(T2, Sw, SwTermPm)
Select Case EQ_NE

Case "EQ": If IsEqStrOpt(R1, R2) Then EvlT1T2 = SomTrue
Case "NE": If IsEqStrOpt(R1, R2) Then EvlT1T2 = SomFalse
Case Else: Thw CSub, "[EQ_NE] does not eq EQ or NE", EQ_NE
End Select
End Function

Private Function EvlT2(T2$, Sw As Dictionary, SwTermPm As Dictionary) As StrOpt
'Return True is evalulated
'switch-term begins with @ or ? or it is *Blank.  @ is for parameter & ? is for switch
'  If @, it will evaluated to str by lookup from Pm
'        if not Has in {Pm}, stop, it means the validation fail to remove this term
'  If ?, it will evaluated to bool by lookup from Sw
'        if not Has in {Sw}, return None
'  Otherwise, just return SomVar(A)
Dim R As StrOpt: R = EvlTerm(T2, Sw, SwTermPm): If Not R.Som Then Exit Function
If FstChr(R.Str) = "*" Then
    If UCase(R.Str) <> "*BLANK" Then Stop ' it means the validation fail to remove this term
    EvlT2 = SomStr("")
Else
    EvlT2 = R
End If
End Function

Private Function EvlTerm(SwTerm$, Sw As Dictionary, SwTermPm As Dictionary) As StrOpt
Dim O$
Select Case True
Case SwTermPm.Exists(SwTerm): O = SwTermPm(SwTerm)
Case Not Sw.Exists(SwTerm):       Exit Function
Case Else:                    O = Sw(SwTerm)
End Select
EvlTerm = SomStr(O)
End Function

Private Function EvlTermAy1(SwTermAy$(), AND_OR$, Sw As Dictionary, BoolTermPm As Dictionary) As Boolean()
Dim R As BoolOpt, O() As Boolean, I
For Each I In SwTermAy
    R = EvlBoolTerm(I, Sw, BoolTermPm)
    If Not R.Som Then Exit Function
    PushI O, R.Bool
Next
EvlTermAy1 = O
End Function

Private Function EvlTermAy(SwTermAy$(), AND_OR$, Sw As Dictionary, BoolTermPm As Dictionary) As BoolOpt
Dim O As Boolean
If Si(SwTermAy) = 0 Then Stop
Dim BoolAy() As Boolean
    BoolAy = EvlTermAy1(SwTermAy, AND_OR, Sw, BoolTermPm)
    If Si(BoolAy) = 0 Then Exit Function
    
Select Case AND_OR
Case "AND": O = IsAllTrue(BoolAy)
Case "OR":  O = IsSomTrue(BoolAy)
Case Else: Stop
End Select
EvlTermAy = SomBool(O)
End Function

Private Function SwBrkAyFmt(A() As SwBrk) As String()
Dim I
For Each I In Itr(A)
    PushI SwBrkAyFmt, CvSwBrk(I).Lin
Next
End Function

Private Property Get OpStrAy() As String()
Static X$()
If Si(X) = 0 Then X = SySsl("OR AND NE EQ")
OpStrAy = X
End Property

Private Sub BrwSwBrkAy(A() As SwBrk)
BrwAy SwBrkAyFmt(A)
End Sub

Private Sub Z_SwRslt()
Dim Stmt As Dictionary, Fld As Dictionary, Er$()
Dim R As SwRslt
R = SwRsltzLnxAy(Samp.SwLnxAy, Samp.Pm)
Brw LyzNNAp("SwLnxAy Pm StmtSw FldSw", _
    ErzLnxAyT1ss(Samp.SwLnxAy, ""), _
    Samp.Pm, _
    R.StmtSw, _
    R.FldSw)
End Sub

Private Sub ZZ()
Dim A() As Lnx
Dim B As Dictionary
Dim C() As SwBrk
Dim XX
End Sub

Private Sub Z()
Z_SwRslt
End Sub

Function CvSwBrk(A) As SwBrk
Set CvSwBrk = A
End Function
