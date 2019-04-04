Attribute VB_Name = "MVb_Str_SubStr"
Option Explicit

Function LasChr$(A)
LasChr = Right(A, 1)
End Function
Function SndChr$(A)
SndChr = Mid(A, 2, 1)
End Function
Function FstAsc%(A)
FstAsc = Asc(FstChr(A))
End Function
Function SndAsc%(A)
SndAsc = Asc(SndChr(A))
End Function
Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function FstTwoChr$(A)
FstTwoChr = Left(A, 2)
End Function

Function SubStrCnt&(S, SubStr)
Dim P&: P = 1
Dim O&, L%
L = Len(SubStr)
While P > 0
    P = InStr(P, S, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Function
    O = O + 1
    P = P + L
Wend
End Function

Function PoszSubStr(A, SubStr$) As Pos
Dim P&: P = InStr(A, SubStr)
If P = 0 Then Exit Function
PoszSubStr = Pos(P, P + Len(SubStr) - 1)
End Function

Private Sub Z_SubStrCnt()
Dim A$, SubStr$
A = "aaaa":                 SubStr = "aa":  Ept = CLng(2): GoSub Tst
A = "aaaa":                 SubStr = "a":   Ept = CLng(4): GoSub Tst
A = "skfdj skldfskldf df ": SubStr = " ":   Ept = CLng(3): GoSub Tst
Exit Sub
Tst:
    Act = SubStrCnt(A, SubStr)
    C
    Return
End Sub

Function DotCnt&(S)
DotCnt = SubStrCnt(S, ".")
End Function

