Attribute VB_Name = "MDao_Lid_Mis_Fun"
Option Explicit

Private Function FnyEptFb(Fbn$, B As LiFb) As String()
Dim J%
'For J = 0 To UBound(B.Ay)
'    If B.Ay(J).Fbn = Fbn Then
'        FnyEptFb = B.Ay(J).Fny
'    End If
'Next
Thw CSub, "Fbn should always be found in LiActFb", "Fbn", Fbn
End Function

Private Function ExtNyFxnLiFx(Fxn, B As LiFx) As String()
Dim J%
'For J = 0 To UBound(B.Ay)
'    If B.Ay(J).Fxn = Fxn Then
'        ExtNyFxnLiFx = ExtNyLiFxcAy(B.Ay(J).FxcAy)
'    End If
'Next
End Function

Private Function EptShtTyAy(Ept() As LiFxc, ExtNm$) As String()
Dim J%
For J = 0 To UBound(Ept)
'    If Ept(J).ExtNm = ExtNm Then
'        EptShtTyAy = CmlAy(Ept(J).ShtTyLis)
        Exit Function
'    End If
Next
End Function

Private Function EptFxcAy(A() As LiFx, Fxn) As LiFxc()
Dim J%
For J = 0 To UBound(A)
    If A(J).Fxn = Fxn Then
'        EptFxcAy = A(J).FxcAy
        Exit Function
    End If
Next
End Function
