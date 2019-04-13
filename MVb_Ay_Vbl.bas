Attribute VB_Name = "MVb_Ay_Vbl"
Option Explicit

Function SyzVbl(Vbl) As String()
SyzVbl = SplitVBar(Vbl)
End Function
Function ItrVbl(Vbl)
ItrVbl = Itr(SyzVbl(Vbl))
End Function

Function LineszVbl$(Vbl$)
LineszVbl = Replace(Vbl, vbCrLf, "|")
End Function

Function IsVbl(A$) As Boolean
Select Case True
Case Not IsStr(A)
Case HasSubStr(A, vbCr)
Case HasSubStr(A, vbLf)
Case Else: IsVbl = True
End Select
End Function

Function IsVblAy(VblAy$()) As Boolean
If Si(VblAy) = 0 Then IsVblAy = True: Exit Function
Dim Vbl
For Each Vbl In VblAy
    If Not IsVbl(CStr(Vbl)) Then Exit Function
Next
IsVblAy = True
End Function

Function IsVdtVbl(Vbl$) As Boolean
If HasSubStr(Vbl, vbCr) Then Exit Function
If HasSubStr(Vbl, vbLf) Then Exit Function
IsVdtVbl = True
End Function

