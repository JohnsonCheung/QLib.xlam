Attribute VB_Name = "MVb_Str_Quote"
Option Explicit
Function BrkQuote(QuoteStr) As S1S2
Dim L%: L = Len(QuoteStr)
Dim s1$, s2$
Select Case L
Case 0:
Case 1
    s1 = QuoteStr
    s2 = QuoteStr
Case 2
    s1 = Left(QuoteStr, 1)
    s2 = Right(QuoteStr, 1)
Case Else
    If InStr(QuoteStr, "*") > 0 Then
        Set BrkQuote = Brk(QuoteStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
Set BrkQuote = S1S2(s1, s2)
End Function

Function QuoteBkt$(A)
QuoteBkt = "(" & A & ")"
End Function
Function QuoteDot$(S)
QuoteDot = "." & S & "."
End Function
Function Quote$(A, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .s1 & A & .s2
End With
End Function

Function QuoteDblVb$(A)
QuoteDblVb = QuoteDbl(Replace(A, vbDblQuote, vbTwoDblQuote))
End Function

Function QuoteDbl$(A)
QuoteDbl = vbDblQuote & A & vbDblQuote
End Function

Function QuoteSng$(A)
QuoteSng = "'" & A & "'"
End Function
Function QuoteSq$(A)
QuoteSq = "[" & A & "]"
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

