VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampS1S2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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

Property Get S1S2AyzLines() As S1S2()
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushObj S1S2AyzLines, S1S2(LineszVbl(A1(J)), LineszVbl(A2(J)))
Next
End Property

Property Get S1S2AyzLin() As S1S2()
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushObj S1S2AyzLin, S1S2(A1(J), A2(J))
Next
End Property

