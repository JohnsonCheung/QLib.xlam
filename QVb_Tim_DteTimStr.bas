Attribute VB_Name = "QVb_Tim_DteTimStr"
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
Property Get NowDteTimStr$()
NowDteTimStr = DteTimStr(Now)
End Property

Property Get NowStr$()
NowStr = NowDteTimStr
End Property


