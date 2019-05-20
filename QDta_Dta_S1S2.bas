Attribute VB_Name = "QDta_Dta_S1S2"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_S1S2."
Private Const Asm$ = "QDta"
Function DrszS1S2s(A As S1S2s) As Drs
DrszS1S2s = DrszFF("S1 S2", S1S2sDry(A))
End Function
Function AvzS1S2(A As S1S2) As Variant()
AvzS1S2 = Array(A.S1, A.S2)
End Function
Function SyzS1S2(A As S1S2) As String()
SyzS1S2 = Sy(A.S1, A.S2)
End Function
Function S1S2sDry(A As S1S2s) As Variant()
Dim J&
For J = 0 To A.N - 1
    PushI S1S2sDry, SyzS1S2(A.Ay(J))
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

Property Get SampS1S2zwLines() As S1S2s
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushS1S2 SampS1S2zwLines, S1S2(LineszVbl(A1(J)), LineszVbl(A2(J)))
Next
End Property

Property Get SampS1S2szwLin() As S1S2s
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushS1S2 SampS1S2szwLin, S1S2(A1(J), A2(J))
Next
End Property


