Attribute VB_Name = "QVb_Str_Quote"
Option Explicit
Private Const CMod$ = "MVb_Str_Quote."
Private Const Asm$ = "QVb"
Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim S1$, S2$
Select Case L
Case 0:
Case 1
    S1 = QuoteStr
    S2 = QuoteStr
Case 2
    S1 = Left(QuoteStr, 1)
    S2 = Right(QuoteStr, 1)
Case Else
    If InStr(QuoteStr, "*") > 0 Then
        BrkQuote = Brk(QuoteStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
BrkQuote = S1S2(S1, S2)
End Function
Sub AsgQuote(OQ1$, OQ2$, QuoteStr$)
With BrkQuote(QuoteStr)
    OQ1 = .S1
    OQ2 = .S2
End With
End Sub
Function QuoteBigBkt$(S)
QuoteBigBkt = "{" & S & "}"
End Function

Function QuoteBkt$(S)
QuoteBkt = "(" & S & ")"
End Function
Function QuoteDot$(S)
QuoteDot = "." & S & "."
End Function
Function Quote$(S, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & S & .S2
End With
End Function

Function QuoteDblVb$(S)
QuoteDblVb = QuoteDbl(Replace(S, vbDblQuote, vbTwoDblQuote))
End Function

Function QuoteDbl$(S)
QuoteDbl = vbDblQuote & S & vbDblQuote
End Function

Function QuoteSng$(S)
QuoteSng = "'" & S & "'"
End Function
Function QuoteSq$(S)
QuoteSq = "[" & S & "]"
End Function
Function QuoteSqIf$(S)
If IsNeedQuote(S) Then QuoteSqIf = QuoteSq(S) Else QuoteSqIf = S
End Function
Function QuoteSqAv(Av()) As String()
Dim I
For Each I In Av
    PushI QuoteSqAv, QuoteSq(I)
Next
End Function

