VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Fx() As LiFx
Private A_Fb() As LiFb
Private A_Fil() As LiFil
Public Apn$
Friend Function Init(Apn$, Fil() As LiFil, Fx() As LiFx, Fb() As LiFb) As LiPm
Me.Apn = Apn
A_Fil = Fil
A_Fx = Fx
A_Fb = Fb
Set Init = Me
End Function
Property Get Fil() As LiFil()
Fil = A_Fil
End Property
Property Get Fx() As LiFx()
Fx = A_Fx
End Property
Property Get Fb() As LiFb()
Fb = A_Fb
End Property

Property Get FfnAy() As String()
Dim J%
For J = 0 To UB(A_Fil)
    With A_Fil(J)
        PushI FfnAy, .Ffn
    End With
Next
End Property
Property Get ExistFfn() As Aset
Set ExistFfn = ExistFfnAset(FfnAy)
End Property
Property Get MisFfn() As Aset
Set MisFfn = MisFfnAset(FfnAy)
End Property
Property Get ExistFilNmToFfnDic() As Dictionary
Dim J%
Set ExistFilNmToFfnDic = New Dictionary
For J = 0 To UB(A_Fil)
    With A_Fil(J)
        If ExistFfn.Has(.Ffn) Then
            ExistFilNmToFfnDic.Add .FilNm, .Ffn
        End If
    End With
Next
End Property

Property Get FmtFil() As String()
Dim J%
For J = 0 To UB(A_Fil)
    With A_Fil(J)
        PushI FmtFil, "FilNm Ffn: " & .FilNm & " " & .Ffn
    End With
Next
End Property


Function FilNmToFfnDic() As Dictionary
Dim J%
Set FilNmToFfnDic = New Dictionary
For J = 0 To UB(A_Fil)
    With A_Fil(J)
        FilNmToFfnDic.Add .FilNm, .Ffn
    End With
Next
End Function

Sub Brw()
Ds.Brw
End Sub

Function Ds() As Ds
Set Ds = DszLidPm(Me)
End Function

