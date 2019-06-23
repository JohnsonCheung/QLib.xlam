Attribute VB_Name = "QVb_Tim_DteTimStr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Tim_DteTimStr."
Private Const Asm$ = "QVb"
Function IsDteTimStr(Str) As Boolean
If Len(Str) <> 17 Then Exit Function
Select Case True
Case IsHHMMSS(Right(Str, 6)), IsYYYYDashMMDashMM(Left(Str, 10)): IsDteTimStr = True
End Select
End Function
Function IsHHMMSS(HHMMSS$) As Boolean
On Error GoTo X
Dim T As Date: T = CDate(Format(HHMMSS, "00:00:00"))
IsHHMMSS = Format(T, "HHMMSS")
Exit Function
X:
End Function
Function IsYYYYDashMMDashMM(A$) As Boolean
Select Case True
Case Len(A) <> 10, Mid(A, 5, 1) <> "-", Mid(A, 8, 1) <> "-": Exit Function
End Select
On Error GoTo X
Dim T As Date: T = CDate(A)
IsYYYYDashMMDashMM = Format(T, "YYYY-MM-DD")
Exit Function
X:
End Function
Function TimStr$(A As Date)
TimStr = DteTimStr(A)
End Function

Function DteTimStr$(A As Date)
DteTimStr = Format(A, "YYYY-MM-DD HHMMSS")
End Function
Function DteTimId$(A As Date)
DteTimId = Format(A, "YYYY_MM_DD_HHMMSS")
End Function
Function NowId$()
NowId = DteTimId(Now)
End Function

Property Get NowStr$()
NowStr = DteTimStr(Now)
End Property



Function CvDbl(S, Optional Fun$)
'Ret : a dbl of @S if can be converted, otherwise empty and debug.print S$
On Error GoTo X
CvDbl = CDbl(S)
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dbl, emp is ret"
End Function

