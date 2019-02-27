Attribute VB_Name = "MVb_Str_Bkt"
Option Explicit
Const CMod$ = "MVbStrBkt."

Private Sub Z_AsgBktPos()
Dim A$, EptFmPos%, EptToPos%
'
A = "(A(B)A)A"
EptFmPos = 1
EptToPos = 7
GoSub Tst
'
A = " (A(B)A)A"
EptFmPos = 2
EptToPos = 8
GoSub Tst
'
A = " (A(B)A )A"
EptFmPos = 2
EptToPos = 9
GoSub Tst
'
Exit Sub
Tst:
    Dim ActFmPos%, ActToPos%
    AsgBktPos A, "(", ActFmPos, ActToPos
    Ass IsEq(ActFmPos, EptFmPos)
    Ass IsEq(ActToPos, EptToPos)
    Return
End Sub

Private Sub Z_Brk_Bkt()
Dim A$, OpnBkt$
A = "aaaa((a),(b))xxx":    OpnBkt = "(":          Ept = Sy("aaaa", "(a),(b)", "xxx"): GoSub Tst
Exit Sub
Tst:
    Act = BrkBkt(A, OpnBkt)
    C
    Return
End Sub

Function IsEqStr(A, B, Optional IgnoreCase As Boolean) As Boolean
IsEqStr = StrComp(A, B, IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)) = 0
End Function

Sub AsgBktPos(A, OpnBkt$, OFmPos%, OToPos%)
Const CSub$ = CMod & "AsgBktPos"
OFmPos = 0
OToPos = 0
'-- OFmPos
    Dim Q1$, Q2$
    Q1 = OpnBkt
    Q2 = ClsBkt(OpnBkt)

    OFmPos = InStr(A, Q1)
    If OFmPos = 0 Then Exit Sub
'-- OToPos
    Dim NOpn%, J%
    For J = OFmPos + 1 To Len(A)
        Select Case Mid(A, J, 1)
        Case Q2
            If NOpn = 0 Then
                OToPos = J
                Exit For
            End If
            NOpn = NOpn - 1
        Case Q1
            NOpn = NOpn + 1
        End Select
    Next
    If OToPos = 0 Then Thw CSub, "The bracket-[Q1]-[Q2] in [Str] has is not in pair: [Q1-Pos] is found, but not Q2-pos is 0", Q1, Q2, A, OFmPos
End Sub

Function ClsBkt$(OpnBkt$)
Select Case OpnBkt
Case "(": ClsBkt = ")"
Case "[": ClsBkt = "]"
Case "{": ClsBkt = "}"
Case Else: Stop
End Select
End Function

Function BrkBkt(A, Optional OpnBkt$ = vbOpnBkt) As String()
Dim P1%, P2%
    AsgBktPos A, OpnBkt, _
    P1, P2
Dim A1$, A2$, A3$
A1 = Left(A, P1 - 1)
A2 = Mid(A, P1 + 1, P2 - P1 - 1)
A3 = Mid(A, P2 + 1)
BrkBkt = Sy(A1, A2, A3)
End Function

Function TakBetBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
AsgBktPos A, OpnBkt, P1, P2
TakBetBkt = Mid(A, P1 + 1, P2 - P1 - 1)
End Function

Function TakAftBkt$(Lin, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
AsgBktPos Lin, OpnBkt, P1, P2
If P2 = 0 Then Exit Function
TakAftBkt = Mid(Lin, P2 + 1)
End Function

Function TakBefBkt$(Lin, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
   AsgBktPos Lin, OpnBkt, P1, P2
If P1 = 0 Then Exit Function
TakBefBkt = Left(Lin, P1 - 1)
End Function

Private Sub Z()
Z_Brk_Bkt
Z_AsgBktPos
MVb_Str_Bkt:
End Sub
