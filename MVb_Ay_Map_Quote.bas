Attribute VB_Name = "MVb_Ay_Map_Quote"
Option Explicit
Function AyQuote(A, QuoteStr$) As String()
If Sz(A) = 0 Then Exit Function
Dim U&: U = UB(A)
Dim O$()
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    With BrkQuote(QuoteStr)
        Q1 = .S1
        Q2 = .S2
    End With
    For J = 0 To U
        O(J) = Q1 & A(J) & Q2
    Next
AyQuote = O
End Function

Function AyQuoteDbl(A) As String()
AyQuoteDbl = AyQuote(A, """")
End Function

Function AyQuoteSng(A) As String()
AyQuoteSng = AyQuote(A, "'")
End Function

Function AyQuoteSq(A) As String()
AyQuoteSq = AyQuote(A, "[]")
End Function

Function AyQuoteSqIf(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI AyQuoteSqIf, QuoteSqIf(I)
Next
End Function


