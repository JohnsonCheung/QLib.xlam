Attribute VB_Name = "MVb_Is_Asc"
Option Explicit

Function IsAscDig(A%) As Boolean
IsAscDig = &H30 <= A And A <= &H39
End Function

Function IsAscPrintablezStrI(S, I) As Boolean
IsAscPrintablezStrI = IsAscPrintable(Asc(Mid(S, I, 1)))
End Function
Function IsAscPrintable(A%) As Boolean
Select Case A
Case 0, 1, 9, 10, 13
Case Else: IsAscPrintable = True
End Select
End Function

Function IsAscDigit(A%) As Boolean
If A < 48 Then Exit Function
If A > 57 Then Exit Function
IsAscDigit = True
End Function

Function IsAscFstNmChr(A%) As Boolean
IsAscFstNmChr = IsAscLetter(A)
End Function

Function IsAscLCase(A%) As Boolean
If A < 97 Then Exit Function
If A > 122 Then Exit Function
IsAscLCase = True
End Function
Function IsAscLetterDig(A%) As Boolean
IsAscLetterDig = True
If IsAscLetter(A) Then Exit Function
If IsAscDig(A) Then Exit Function
IsAscLetterDig = False
End Function
Function IsAscLetter(A%) As Boolean
IsAscLetter = True
If IsAscUCase(A) Then Exit Function
If IsAscLCase(A) Then Exit Function
IsAscLetter = False
End Function

Function IsAscNmChr(A%) As Boolean
IsAscNmChr = True
If IsAscLetter(A) Then Exit Function
If IsAscDig(A) Then Exit Function
IsAscNmChr = A = 95 '_
End Function

Function IsAscUCase(A%) As Boolean
If A < 65 Then Exit Function
If A > 90 Then Exit Function
IsAscUCase = True
End Function
