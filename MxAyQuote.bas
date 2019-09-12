Attribute VB_Name = "MxAyQuote"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyQuote."
Function QteSqBkt$(S)
QteSqBkt = "[" & S & "]"
End Function

Function QteSqBktIfzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI QteSqBktIfzAy, QteSqIf(I)
Next
End Function

Function SyQte(Sy$(), QteStr$) As String()
If Si(Sy) = 0 Then Exit Function
Dim U&: U = UB(Sy)
Dim Q1$, Q2$
    With BrkQte(QteStr)
        Q1 = .S1
        Q2 = .S2
    End With

Dim O$()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Q1 & Sy(J) & Q2
    Next
SyQte = O
End Function

Function SyQteDbl(Sy$()) As String()
SyQteDbl = SyQte(Sy, """")
End Function

Function SyQteSng(Sy$()) As String()
SyQteSng = SyQte(Sy, "'")
End Function

Function SyzQteSq(Sy$()) As String()
SyzQteSq = SyQte(Sy, "[]")
End Function

Function SyzQteSqIf(Sy$()) As String()
Dim I: For Each I In Itr(Sy)
    PushI SyzQteSqIf, QteSqIf(I)
Next
End Function
