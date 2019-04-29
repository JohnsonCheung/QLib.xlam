Attribute VB_Name = "MVb_Ay_Map_Quote"
Option Explicit
Function QuoteSqBkt$(S$)
QuoteSqBkt = "[" & S & "]"
End Function
Function SyQuote(Sy$(), QuoteStr$) As String()
If Si(Sy) = 0 Then Exit Function
Dim U&: U = UB(Sy)
Dim Q1$, Q2$
    With BrkQuote(QuoteStr)
        Q1 = .S1
        Q2 = .S2
    End With

Dim O$()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Q1 & Sy(J) & Q2
    Next
SyQuote = O
End Function

Function SyQuoteDbl(Sy$()) As String()
SyQuoteDbl = SyQuote(Sy, """")
End Function

Function SyQuoteSng(Sy$()) As String()
SyQuoteSng = SyQuote(Sy, "'")
End Function

Function SyQuoteSq(Sy$()) As String()
SyQuoteSq = SyQuote(Sy, "[]")
End Function

Function SyQuoteSqIf(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI SyQuoteSqIf, QuoteSqIf(CStr(I))
Next
End Function


