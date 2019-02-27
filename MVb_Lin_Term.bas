Attribute VB_Name = "MVb_Lin_Term"
Option Explicit
Const CMod$ = "MVb_Lin_Term."
Function RmvTermAy$(Lin, Ay$())
Dim T$, I
T = T1(Lin)
For Each I In Ay
    If I = T Then
        RmvTermAy = LTrim(Mid(LTrim(Lin), Len(T) + 1))
        Exit Function
    End If
Next
RmvTermAy = Lin
End Function
Function TLin$(TermAy$())
TLin = JnTermAy(TermAy)
End Function

Function JnTermAy$(TermAy$())
JnTermAy = JnSpc(AyQuoteSqIf(TermAy))
End Function
Function TermAyzTT(TT) As String()
Const CSub$ = CMod & "TermAyzTT"
Select Case True
Case IsStr(TT): TermAyzTT = TermAy(TT)
Case IsSy(TT): TermAyzTT = TT
Case Else: Thw CSub, "Given TT must be Str or Sy", "TypeName TT", TypeName(TT), TT
End Select
End Function
Function LinzTermAy$(TermAy)
LinzTermAy = JnSpc(AyQuoteSqIf(TermAy))
End Function
Function TermAset(Lin) As Aset
Set TermAset = AsetzAy(TermAy(Lin))
End Function
Function TermAy(Lin) As String()
Dim L$, J%
L = Lin
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushNonBlankStr TermAy, ShfTerm(L)
Wend
End Function

Function ShfT$(O)
ShfT = ShfTerm(O)
End Function

Function ShfX(O, X$) As Boolean
If T1(O) = X Then
    ShfX = True
    O = RmvT1(O)
End If
End Function

Private Function ShfTerm1$(O)
Dim A$
AsgAp BrkBkt(O, "["), A, ShfTerm1, O
End Function

Function ShfTerm$(O)
Dim A$
    A = LTrim(O)
If FstChr(A) = "[" Then ShfTerm = ShfTerm1(O): Exit Function
Dim P%
    P = InStr(A, " ")
If P = 0 Then
    ShfTerm = A
    O = ""
    Exit Function
End If
ShfTerm = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Private Sub Z_ShfT()
Dim O$, OEpt$
O = " S   DFKDF SLDF  "
OEpt = "DFKDF SLDF  "
Ept = "S"
GoSub Tst
'
O = " AA BB "
Ept = "AA"
OEpt = "BB "
GoSub Tst
'
Exit Sub
Tst:
    Act = ShfT(O)
    C
    Ass O = OEpt
    Return
End Sub


Private Sub Z()
Z_ShfT
MVb_Lin_Term:
End Sub
