Attribute VB_Name = "MxSampS12s"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxSampS12s."
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

Function SampS12s_wiLines() As S12s
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushS12 SampS12s_wiLines, S12(LineszVbl(A1(J)), LineszVbl(A2(J)))
Next
End Function

Function SampS12s() As S12s
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushS12 SampS12s, S12(A1(J), A2(J))
Next
End Function