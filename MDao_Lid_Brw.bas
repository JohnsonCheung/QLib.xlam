Attribute VB_Name = "MDao_Lid_Brw"
Option Explicit
Private Sub Z_BrwLidPm()
BrwLidPm RptLidPm
End Sub

Sub BrwLidPm(A As LidPm)
BrwDs DszLidPm(A)
End Sub

Function DszLidPm(A As LidPm) As Ds
Set DszLidPm = Ds(DtAy(FilDt(A.Fil), FxColDt(A.Fx), FbColDt(A.Fb)), "LiPm-" & A.Apn)
End Function

Private Function FbColDt(A() As LidFb) As DT
Dim Dry(), J%
For J = 0 To UB(A)
    With A(J)
        PushI Dry, Array(.Fbn, .T, .Fset.TermLin, .Bexpr)
    End With
Next
Set FbColDt = DT("FbCol", "Fbn T FF Bexpr", Dry)
End Function

Private Function FilDt(A() As LidFil) As DT
Set FilDt = DT("LnkFil", "FilNm Ffn Exist", FilDry(A))
End Function

Private Function FilDry(A() As LidFil) As Variant()
Dim J%
For J = 0 To UB(A)
    With A(J)
        PushI FilDry, Array(.FilNm, .Ffn, HasFfn(.Ffn))
    End With
Next
End Function

Private Function FxColDt(A() As LidFx) As DT
Set FxColDt = DT("FxCol", "Fxn Wsn T ColNm ShtTyLis ExtNm", FxColDry(A))
End Function

Private Function FxColDry(A() As LidFx) As Variant()
Dim Dry(), J%
For J = 0 To UB(A)
    PushIAy FxColDry, FxColDr(A(J))
Next
End Function

Private Function FxColDr(A As LidFx) As Variant()
Dim J%, Fxc() As LidFxc
Fxc = A.Fxc
For J = 0 To UB(Fxc)
    With Fxc(J)
        PushI FxColDr, Array(A.Fxn, A.Wsn, A.T, .ColNm, .ShtTyLis, .ExtNm)
    End With
Next
End Function

