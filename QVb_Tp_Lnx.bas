Attribute VB_Name = "QVb_Tp_Lnx"
Option Compare Text
Option Explicit
Enum EmIxFm 'Ix is always 0.  When present as Lno, Ix=0 will be Lno=1.
    EiFm1   'Or, presenting Ix=0 as 1 for Lno
    EiFm0   'So, presenting Ix=0 as 0  for Ix
End Enum
Type SLnoy
    S As String
    Lnoy() As Long
End Type
Type SLnoys: N As Long: Ay() As SLnoy: End Type
Type SIxy
    S As String
    Ixy() As Long
End Type
Type SIxys: N As Long: Ay() As SIxy: End Type
Type Lnx
    Lin As String
    Ix As Long
End Type
Type Lnxs: N As Long: Ay() As Lnx: End Type
Type Lnxses: N As Long: Ay() As Lnxs: End Type
Type LnxsRslt: Er() As String: Lnxs As Lnxs: End Type
Type SomLnx
    Som As Boolean
    Lnx As Lnx
End Type
Function SomLnx(A As Lnx) As SomLnx
SomLnx.Som = True
SomLnx.Lnx = A
End Function
Function LinAyzLnxs(A As Lnxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LinAyzLnxs, A.Ay(J).Lin
Next
End Function
Function IxyzLnxs(A As Lnxs) As Long()
Dim J&
For J = 0 To A.N - 1
    PushI IxyzLnxs, A.Ay(J).Ix
Next
End Function
Private Sub ZZ_FmtLnxs()
Dim A As Lnxs
GoSub ZZ
Exit Sub
ZZ:
    A = Lnxs(SrczP(CPj))
    Brw FmtLnxs(A)
    Return
End Sub

Function FmtLnx(A As Lnx)
FmtLnx = LnxStr(A)
End Function

Function FmtLnxs(A As Lnxs) As String()
Dim B$(): B = AlignRzAy(IxyzLnxs(A))
FmtLnxs = JnAyab(B, LinAyzLnxs(A), " ")
End Function
Function SngLnx(A As Lnx) As Lnxs
PushLnx SngLnx, A
End Function
Function EmpLnxs() As Lnxs
End Function
Function Lnx(Lin, Ix) As Lnx
If Ix < 0 Then Thw CSub, "Ix cannot be negative", "Ix Lin", Ix, Lin
Lnx.Ix = Ix
Lnx.Lin = Lin
End Function

Sub PushLnx(O As Lnxs, M As Lnx)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function LnxszVbl(Vbl$) As Lnxs
LnxszVbl = Lnxs(LyzVbl(Vbl))
End Function
Function Lnxs(Ly$()) As Lnxs
Dim J&
For J = 0 To UB(Ly)
    PushLnx Lnxs, Lnx(Ly(J), J)
Next
End Function
Function IxywLin(A As Lnxs, Lin) As Long()
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
    If .Lin = Lin Then PushI IxywLin, .Ix
    End With
Next
End Function

Function Lnoss$(Ixy() As Long)
Lnoss = JnSpc(AyIncEle1(Ixy))
End Function

Function LnosswLin(A As Lnxs, Lin)
LnosswLin = Lnoss(IxywLin(A, Lin))
End Function
Function LinyzLnxs(A As Lnxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LinyzLnxs, A.Ay(J).Lin
Next
End Function
Function SIxyszDup(A As Lnxs) As SIxys
Dim Dup$, I
For Each I In Itr(AywDup(LinyzLnxs(A)))
    Dup = I
    PushIIxy SIxyszDup, SIxy(Dup, IxyzLnxsLin(A, Dup))
Next
End Function

Function BrwLnxs(A As Lnxs)
B FmtLnxsWiLno(A)
End Function

Sub BrwSIxys(A As SIxys)
B LyzSIxys(A)
End Sub
Function LyzSIxys(A As SIxys, Optional SLnossMacros$) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzSIxys, LinzSIxy(A.Ay(J))
Next
End Function

Function LinzSIxy$(A As SIxy, Optional SLnossMacro$)
Dim M$: M = DftStr(SLnossMacro, "S({S}) Lnoss({Lnoss})")
With A
LinzSIxy = FmtMacro(M, .S, Lnoss(.Ixy))
End With
End Function
Sub PushIIxy(O As SIxys, M As SIxy)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Private Sub Z_DupT1zLnxs()
Dim A As Lnxs, Act As SIxys, Ept As SIxys
GoSub T0
'GoSub ZZ
Exit Sub
ZZ:
    A = Lnxs(SrcV)
    BrwSIxys DupT1zLnxs(A)
    Return
T0:
    A = Lnxs(Sy("A B", " B X", "A", "B", "C"))
    PushIIxy Ept, SIxy("A", LngAp(1, 3))
    PushIIxy Ept, SIxy("B", LngAp(2, 4))
    GoTo Tst
Tst:
    Act = DupT1zLnxs(A)
    BrwSIxys Act
    BrwSIxys Ept
    Debug.Assert IsEqSIxys(Act, Ept)
    Return
End Sub
Function IsEqSIxys(A As SIxys, B As SIxys) As Boolean
With A
If .N <> B.N Then Exit Function
Dim J&
For J = 0 To .N - 1
    If Not IsEqSIxy(.Ay(J), B.Ay(J)) Then Exit Function
Next
End With
IsEqSIxys = True
End Function
Function IsEqSIxy(A As SIxy, B As SIxy) As Boolean
With A
Select Case True
Case .S <> B.S, Not IsEqAy(.Ixy, B.Ixy)
Case Else: IsEqSIxy = True
End Select
End With
End Function
Function DupT1zLnxs(A As Lnxs) As SIxys
DupT1zLnxs = DupLinzLnxs(T1Lnxs(A))
End Function
Function DupLinzLnxs(A As Lnxs) As SIxys
Dim Dup$, I
For Each I In Itr(AywDup(LinyzLnxs(A)))
    Dup = I
    PushIIxy DupLinzLnxs, SIxy(Dup, IxyzLnxsLin(A, Dup))
Next
End Function
Function T1Lnxs(A As Lnxs) As Lnxs 'Take the T1 of A().Lin to return Lnxs
Const Insp As Boolean = True
If Insp Then BrwLnxs A
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
    PushLnx T1Lnxs, Lnx(T1(.Lin), .Ix)
    End With
Next
If Insp Then
    BrwLnxs T1Lnxs
    Stop
End If
End Function

Function IncLnxs(A As Lnxs, Optional By% = 1) As Lnxs
Dim O As Lnxs: O = A
Dim J&
For J = 0 To A.N - 1
    With O.Ay(J): .Ix = .Ix + 1: End With
Next
End Function

Private Function MaxIx%(A As Lnxs)
Dim J&, O&
For J = 0 To A.N - 1
    O = Max(A.Ay(J).Ix, O)
Next
MaxIx = O
End Function

Function IxyzLnxsLin(A As Lnxs, Lin) As Long()
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
        If .Lin = Lin Then PushI IxyzLnxsLin, .Ix
    End With
Next
End Function

Function SIxy(S, Ixy() As Long) As SIxy
SIxy.S = S
SIxy.Ixy = Ixy
End Function

Function LnxsRslt(Lnxs As Lnxs, Er$()) As LnxsRslt
LnxsRslt.Er = Er
LnxsRslt.Lnxs = Lnxs
End Function
Function LnxswUniqT1(A As Lnxs) As Lnxs 'If Lin has T1 dup, take the first one
Dim Dup$(), T1$, J%
For J = 0 To A.N - 1
    With A.Ay(J)
        T1 = T1zS(.Lin)
        If Not HasEle(Dup, T1) Then
            PushI Dup, T1
            PushLnx LnxswUniqT1, A.Ay(J)
        End If
    End With
Next
End Function
Private Function T1AyzLnxs(A As Lnxs) As String()
T1AyzLnxs = T1Ay(LinAyzLnxs(A))
End Function

Function DupT1AyzLnxs(A As Lnxs) As String()
DupT1AyzLnxs = AywDup(T1AyzLnxs(A))
End Function
Private Function ErOfDupT1zLnxs(A As Lnxs) As String()
Dim Dup
For Each Dup In Itr(DupT1AyzLnxs(A))
    PushI ErOfDupT1zLnxs, ErOfDupT1zLnxsDup(A, Dup)
Next
End Function
Private Function FstT1Lno&(A As Lnxs, T1) 'Return the Lno of the first ele in A with T1 as given
Dim J&
For J = 0 To A.N - 1
    If T1zS(A.Ay(J).Lin) = T1 Then FstT1Lno = A.Ay(J).Ix + 1: Exit Function
Next
End Function

Function LnxswT1(A As Lnxs, T1) As Lnxs
Dim J&
For J = 0 To A.N - 1
    If T1zS(A.Ay(J).Lin) = T1 Then PushLnx LnxswT1, A.Ay(J)
Next
End Function
Function RmvFstElezLnxs(A As Lnxs) As Lnxs
Dim J&
For J = 1 To A.N - 1
    PushLnx RmvFstElezLnxs, A.Ay(J)
Next
End Function
Function LnoAyzLnxs(A As Lnxs) As Long()
LnoAyzLnxs = IncAy(IxyzLnxs(A))
End Function
Function IxlinAyzLnxs(A As Lnxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI IxlinAyzLnxs, Ixlin(A.Ay(J))
Next
End Function
Function Ixlin$(A As Lnx)
Ixlin = A.Ix & " " & A.Lin
End Function
Function LnxszIxlinAy(IxlinAy$()) As Lnxs
Dim Ixlin
For Each Ixlin In Itr(IxlinAy)
    PushLnx LnxszIxlinAy, LnxzIxlin(Ixlin)
Next
End Function
Function LnxzIxlin(Ixlin) As Lnx
With BrkSpc(Ixlin)
LnxzIxlin = Lnx(.S2, .S1)
End With
End Function
Function LnxszIxlinVbl(IxlinVbl$) As Lnxs
LnxszIxlinVbl = LnxszIxlinAy(SplitVBar(IxlinVbl))
End Function
Function LnxzStr(LnxStr$) As Lnx
With BrkSpc(LnxStr): LnxzStr = Lnx(.S2, .S1): End With
End Function

Private Function ErOfDupT1zLnxsDup$(A As Lnxs, Dup)
Dim Ix, M$
Dim FstLno&: FstLno = FstT1Lno(A, Dup) ' FstLno has Dup in A
Dim Lno
For Each Lno In LnoAyzLnxs(RmvFstElezLnxs(LnxswT1(A, Dup)))
    M = FmtQQ("Lno[?] has T1[?] already found in Lno[?].  This line is skipped.", Ix, Dup, FstLno)
    PushI ErOfDupT1zLnxsDup, M
Next
End Function
Function LnxswSngT1(A As Lnxs) As Lnxs

End Function
Function LnxsRsltOfDupT1(A As Lnxs) As LnxsRslt
LnxsRsltOfDupT1 = LnxsRslt(LnxswSngT1(A), DupT1ErzLnxs(A))
End Function
Function DupT1ErzLnxs(A As Lnxs) As String()

End Function
Function DupT2AyzLnxs(A As Lnxs) As String()
DupT2AyzLnxs = AywDup(T2Ay(LyzLnxs(A)))
End Function

Function LnxswT1Ay(A As Lnxs, T1Ay$()) As Lnxs
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
    If Not HasEle(T1Ay, T1(.Lin)) Then PushLnx LnxswT1Ay, A.Ay(J)
    End With
Next
End Function

Function FmtLnxsWiLno(A As Lnxs) As String()
Dim J&, O$()
For J = 0 To A.N - 1
    With A.Ay(J)
    PushI O, FmtQQ("Lno#?:[?]", .Ix, .Lin)
    End With
Next
FmtLnxsWiLno = O ' AlignzBySepss(O, ":")
End Function

Function ErzLnxsT1ss(A As Lnxs, T1ss$) As String()
Dim T1Ay$(): T1Ay = SyzSS(T1ss)
If Si(T1Ay) = 0 Then Exit Function
Dim Er As Lnxs: ' Er = LnxSyeT1Sy(A, T1Ay)
'If Si(Er) = 0 Then Exit Function
'ErzLnxsT1ss = LyzMsgNap("There are lines have invalid T1", "Lines Valid-Ty", LyzLnxszWithLno(Er), T1Ay)
End Function

Function LnxswT2(A As Lnxs, T2) As Lnxs
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
        If T2zS(.Lin) = T2 Then
            PushLnx LnxswT2, A.Ay(J)
        End If
    End With
Next
End Function

Function LnxswRmvgT1(A As Lnxs, T) As Lnxs
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
        If T1(.Lin) = T Then
            PushLnx LnxswRmvgT1, Lnx(RmvT1(.Lin), .Ix)
        End If
    End With
Next
End Function

Function RmvT1zLnx$(A As Lnx)
RmvT1zLnx = RmvT1(A.Lin)
End Function

Function LnxStr$(A As Lnx)
LnxStr = A.Ix & " " & A.Lin
End Function

Function LinzLnx$(A As Lnx)
LinzLnx = LnxStr(A)
End Function

Function LyzLnxs(A As Lnxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzLnxs, LinzLnx(A.Ay(J))
Next
End Function
Private Function Lnoss_FmLnxs_WhT1$(A As Lnxs, T1)
Dim J%, O&()
For J = 0 To A.N - 1
    If T1zS(A.Ay(J).Lin) = T1 Then
        PushI O, A.Ay(J).Ix + 1
    End If
Next
Lnoss_FmLnxs_WhT1 = JnSpc(O)
End Function
Private Function LnossAy_FmLnxs_WhT1Ay(A As Lnxs, T1Ay$()) As String()
Dim T1
For Each T1 In Itr(T1Ay)
    PushI LnossAy_FmLnxs_WhT1Ay, Lnoss_FmLnxs_WhT1(A, T1)
Next
End Function
Private Sub ZZ_DupT1_FmLnxs_ToDupT1Ay_AndLnossAy()
Dim EptDupT1Ay$(), EptLnossAy$()
Dim ActDupT1Ay$(), ActLnossAy$()
Dim A As Lnxs
GoSub ZZ
Exit Sub
ZZ:
    A = LnxszVbl("a b c|a b|b 1|a 1")
    DupT1_FmLnxs_ToDupT1Ay_AndLnossAy A, ActDupT1Ay, ActLnossAy
    D "---"
    D ActDupT1Ay
    D "---"
    D ActLnossAy
    Return
End Sub

Sub DupT1_FmLnxs_ToDupT1Ay_AndLnossAy(A As Lnxs, ODupT1Ay$(), OLnossAy$())
ODupT1Ay = DupT1AyzLnxs(A)
OLnossAy = LnossAy_FmLnxs_WhT1Ay(A, ODupT1Ay)
End Sub
