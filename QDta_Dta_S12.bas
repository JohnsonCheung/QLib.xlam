Attribute VB_Name = "QDta_Dta_S12"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_S12."
Private Const Asm$ = "QDta"
Function DrszS12s(A As S12s) As Drs
DrszS12s = DrszFF("S1 S2", DyzS12s(A))
End Function
Function AvzS12(A As S12) As Variant()
AvzS12 = Array(A.S1, A.S2)
End Function
Function SyzS12(A As S12) As String()
SyzS12 = Sy(A.S1, A.S2)
End Function
Function DyzS12s(A As S12s) As Variant()
Dim J&
For J = 0 To A.N - 1
    PushI DyzS12s, SyzS12(A.Ay(J))
Next
End Function

Private Sub X(O1$(), O2$())
Erase O1
Erase O2
Dim A1$, A2$
A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df":          GoSub X
A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub X
A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df":               GoSub X
A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df":           GoSub X
A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df":            GoSub X
Exit Sub
X:
    PushI O1, LineszVbl(A1)
    PushI O2, LineszVbl(A2)
    Return
End Sub

Property Get SampS12zwLines() As S12s
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushS12 SampS12zwLines, S12(LineszVbl(A1(J)), LineszVbl(A2(J)))
Next
End Property

Property Get SampS12szwLin() As S12s
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushS12 SampS12szwLin, S12(A1(J), A2(J))
Next
End Property


