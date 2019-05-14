Attribute VB_Name = "QVb_Ay_Map_Quote"
Option Explicit
Private Const CMod$ = "MVb_Ay_Map_Quote."
Private Const Asm$ = "QVb"
Function QuoteSqBkt$(S)
QuoteSqBkt = "[" & S & "]"
End Function
Function QuoteSqBktIfzSy$(Sy$())
Dim I
For Each I In Itr(Sy)
    PushI QuoteSqBktIfzSy, QuoteSqIf(CStr(I))
Next
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

Function QuoteSqzAy(Sy$()) As String()
QuoteSqzAy = SyQuote(Sy, "[]")
End Function

Function QuoteSqzAyIf(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI QuoteSqzAyIf, QuoteSqIf(CStr(I))
Next
End Function


