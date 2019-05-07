Attribute VB_Name = "QVb_JnSplit_Jn"
Option Explicit
Private Const CMod$ = "MVb_JnSplit_Jn."
Private Const Asm$ = "QVb"
Function Jn$(Ay, Optional Sep$ = "")
Jn = Join(SyzAy(Ay), Sep)
End Function
Function QuoteBktJnComma$(Ay)
QuoteBktJnComma = QuoteBkt(JnComma(Ay))
End Function

Function JnComma$(Ay)
JnComma = Jn(Ay, ",")
End Function

Function JnBq$(Ay)
JnBq = Jn(Ay, "`")
End Function

Function JnCommaCrLf$(Ay)
JnCommaCrLf = Jn(Ay, "," & vbCrLf)
End Function

Function JnAnd$(Ay)
JnAnd = Jn(Ay, " and ")
End Function

Function JnCommaSpc$(Ay)
JnCommaSpc = Jn(Ay, ", ")
End Function

Function JnCrLf$(Ay)
JnCrLf = Jn(Ay, vbCrLf)
End Function

Function JnDblCrLf$(Ay)
JnDblCrLf = Jn(Ay, vbCrLf & vbCrLf)
End Function

Function JnDotAp$(ParamArray Ap())
Dim Av(): Av = Ap: JnDotAp = JnDot(Av)
End Function
Function JnQDot$(Ay) 'JnQDot = QuoteDot . JnDot
JnQDot = QuoteDot(JnDot(Ay))
End Function

Function JnDot$(Ay)
JnDot = Jn(Ay, ".")
End Function

Function JnDollar$(Ay)
JnDollar = Jn(Ay, "$")
End Function

Function JnDblDollar$(Ay)
JnDblDollar = Jn(Ay, "$$")
End Function

Function JnPthSep$(Ay)
JnPthSep = Jn(Ay, PthSep)
End Function

Function JnQDblComma$(Sy$())
JnQDblComma = JnComma(SyQuoteDbl(Sy))
End Function

Function JnQDblSpc$(Sy$())
JnQDblSpc = JnSpc(SyQuoteDbl(Sy))
End Function

Function JnQSngComma$(Sy$())
JnQSngComma = JnComma(SyQuoteSng(Sy))
End Function

Function JnQSngSpc$(Sy$())
JnQSngSpc = JnSpc(SyQuoteSng(Sy))
End Function

Function JnQSqCommaSpc$(Sy$())
JnQSqCommaSpc = JnCommaSpc(SyQuoteSqIf(Sy))
End Function

Function JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(SyQuoteSq(SyzAy(Ay)))
End Function

Function JnSemi$(Ay)
JnSemi = Jn(Ay, ";")
End Function

Function JnOr$(Ay)
JnOr = Jn(Ay, " or ")
End Function

Function JnSpc$(Ay)
JnSpc = Jn(Ay, " ")
End Function

Function JnTab$(Ay)
JnTab = Join(Ay, vbTab)
End Function

Function JnTerm$(Ay)
Dim O$(), X
For Each X In Itr(Ay)
'    PushI O, QuoteSq(CStr(X))
Next
JnTerm = Join(O, " ")
End Function

Function JnVBar$(Ay)
JnVBar = Jn(Ay, "|")
End Function

Function JnVbarSpc$(Ay)
JnVbarSpc = Jn(Ay, " | ")
End Function

