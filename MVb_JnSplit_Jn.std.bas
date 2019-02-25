Attribute VB_Name = "MVb_JnSplit_Jn"
Option Explicit
Function Jn$(A, Optional Sep$ = "")
Jn = Join(SyzAy(A), Sep)
End Function
Function QuoteBktJnComma$(Ay)
QuoteBktJnComma = QuoteBkt(JnComma(Ay))
End Function

Function JnComma$(A)
JnComma = Jn(A, ",")
End Function

Function JnCommaCrLf$(A)
JnCommaCrLf = Jn(A, "," & vbCrLf)
End Function

Function JnAnd$(A)
JnAnd = Jn(A, " and ")
End Function

Function JnCommaSpc$(A)
JnCommaSpc = Jn(A, ", ")
End Function

Function JnCrLf$(Ay)
JnCrLf = Jn(Ay, vbCrLf)
End Function

Function JnDblCrLf$(A)
JnDblCrLf = Jn(A, vbCrLf & vbCrLf)
End Function

Function JnDot$(A)
JnDot = Jn(A, ".")
End Function

Function JnDollar$(A)
JnDollar = Jn(A, "$")
End Function

Function JnDblDollar$(A)
JnDblDollar = Jn(A, "$$")
End Function

Function JnPthSep$(A)
JnPthSep = Jn(A, PthSep)
End Function

Function JnQDblComma$(A)
JnQDblComma = JnComma(AyQuoteDbl(A))
End Function

Function JnQDblSpc$(A)
JnQDblSpc = JnSpc(AyQuoteDbl(A))
End Function

Function JnQSngComma$(A)
JnQSngComma = JnComma(AyQuoteSng(A))
End Function

Function JnQSngSpc$(A)
JnQSngSpc = JnSpc(AyQuoteSng(A))
End Function

Function JnQSqCommaSpc$(A)
JnQSqCommaSpc = JnCommaSpc(AyQuoteSq(A))
End Function

Function JnQSqBktSpc$(A)
JnQSqBktSpc = JnSpc(AyQuoteSq(SyzAy(A)))
End Function

Function JnSemi$(A)
JnSemi = Jn(A, ";")
End Function

Function JnOr$(A)
JnOr = Jn(A, " or ")
End Function

Function JnSpc$(A)
JnSpc = Jn(A, " ")
End Function

Function JnTab$(A)
JnTab = Join(A, vbTab)
End Function

Function JnTerm$(A)
Dim O$(), X
For Each X In Itr(A)
'    PushI O, QuoteSq(CStr(X))
Next
JnTerm = Join(O, " ")
End Function

Function JnVBar$(A)
JnVBar = Jn(A, "|")
End Function

Function JnVBarSpc$(A)
JnVBarSpc = Jn(A, " | ")
End Function

