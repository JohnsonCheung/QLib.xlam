VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fxn$, Wsn$, T$, Bexpr$
Private A_Fxc() As LiFxc
Friend Function Init(Fxn, Wsn, T, Fxc() As LiFxc, Bexpr) As LiFx
Set Init = Me
With Me
    .Fxn = Fxn
    .Wsn = Wsn
    .T = T
    A_Fxc = Fxc
   .Bexpr = Bexpr
End With
End Function
Property Get FxcAy() As LiFxc()
FxcAy = A_Fxc
End Property

Property Get Fset() As Aset
Set Fset = AsetzAy(Fny)
End Property

Property Get Fny() As String()
Dim J%
For J = 0 To UB(A_Fxc)
    PushI Fny, A_Fxc(J).ColNm
Next
End Property

Function EptFset() As Aset
Set EptFset = AsetzAy(ExtNy)
End Function

Function ExtNy() As String()
Dim J%
For J = 0 To UB(A_Fxc)
    PushI ExtNy, A_Fxc(J).Extnm
Next
End Function

Friend Function ExistFx$(B() As LiActFx)
Dim J%
For J = 0 To UB(B)
    With B(J)
        If Fxn = .Fxn Then
            If Wsn = .Wsn Then
                ExistFx = .Fx
                Exit Function
            End If
        End If
    End With
Next
End Function

