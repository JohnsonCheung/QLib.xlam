VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiFb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fbn$, T$, Fset As Aset, Bexpr$
Friend Function Init(Fbn, T, Fset As Aset, Bexpr) As LiFb
With Me
.Fbn = Fbn
.T = T
Set .Fset = Fset
.Bexpr = Bexpr
End With
Set Init = Me
End Function
Friend Function ExistFb$(A() As LiActFb)
Dim J%
For J = 0 To UB(A)
    With A(J)
        If Fbn = .Fbn Then
            If T = .T Then
                ExistFb = .Fb
                Exit Function
            End If
        End If
    End With
Next
End Function

