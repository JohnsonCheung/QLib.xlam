Attribute VB_Name = "QVb_F_JnSplit_Jn"
Option Compare Text
Option Explicit
Option Base 0
Private Const CMod$ = "MVb_JnSplit_Jn."
Private Const Asm$ = "QVb"
Function Jn$(Ay, Optional Sep$ = "")
Jn = Join(SyzAy(Ay), Sep)
End Function
Function QteBktJnComma$(Ay)
QteBktJnComma = QteBkt(JnComma(Ay))
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

Function CatAy(AyA, AyB, Optional Sep$) As String()

End Function

Function JnAp$(Sep$, ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap
JnAp = Jn(Av, Sep)
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

Function JnCrLfAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnCrLfAp = Jn(Av, vbCrLf)
End Function

Function JnDblCrLf$(Ay)
JnDblCrLf = Jn(Ay, vb2CrLf)
End Function

Function JnDotAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) > 0 Then Av = Ap: JnDotAp = JnDot(Av)
End Function

Function QteJnzAsTLin$(Ay)
QteJnzAsTLin = QteJn(Ay, " | ", "| * |")
End Function

Function QteJn$(Ay, Sep$, QteStr$)
QteJn = Qte(Jn(Ay, Sep), QteStr)
End Function
Function QteJnDot$(Ay)
'Ret : a str joining @Ay and qte with . in front and at end
QteJnDot = QteDot(JnDot(Ay))
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
JnQDblComma = JnComma(SyQteDbl(Sy))
End Function

Function JnQDblSpc$(Sy$())
JnQDblSpc = JnSpc(SyQteDbl(Sy))
End Function

Function JnQSngComma$(Sy$())
JnQSngComma = JnComma(SyQteSng(Sy))
End Function

Function JnQSngSpc$(Sy$())
JnQSngSpc = JnSpc(SyQteSng(Sy))
End Function

Function JnQSqCommaSpc$(Sy$())
JnQSqCommaSpc = JnCommaSpc(SyzQteSqIf(Sy))
End Function

Function JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(SyzQteSq(SyzAy(Ay)))
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
    PushI O, QteSq(X)
Next
JnTerm = Join(O, " ")
End Function

Function JnVBar$(Ay)
JnVBar = Jn(Ay, "|")
End Function

Function JnVbarSpc$(Ay)
JnVbarSpc = Jn(Ay, " | ")
End Function


'
