Attribute VB_Name = "QIde_Ens_EnsMdy"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_MthMdy."
Function MthLinzEnsPrv$(MthLin)
Const CSub$ = CMod & "MthLinzEnsPrv"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", "Lin", MthLin
MthLinzEnsPrv = "Private " & RmvMdy(MthLin)
End Function

Function MthLinzEnsPub$(MthLin)
Const CSub$ = CMod & "MthLinzEnsPub"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", MthLin
MthLinzEnsPub = RmvMdy(MthLin)
End Function

Sub EnsPjPrvZ()
EnsPrvZzP CPj
End Sub

Sub EnsPrvZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsPrvZzM C.CodeModule
Next
End Sub
Private Function Z_EnsPrv(A As Drs) As Drs

End Function
Sub EnsPrvZzM(M As CodeModule, Optional Rpt)
Const CmPfx$ = "X_"
Dim A As Drs: ' A = DPubZMth(M) ' L MthLin
Dim B As Drs: ' B = X_EnsPrv(A)   ' L MthLin PrvZ
Dim C As Drs: C = SelDrszAs(B, "L PrvZ:NewL MthLin:OldL")

RplLin M, C
End Sub

Function LnoAyOfPubZ(M As CodeModule) As Long()
Dim L, J&
For Each L In Itr(Src(M))
    J = J + 1
    'If IsMthLinzPub(L) Then
        PushI LnoAyOfPubZ, J
    'End If
Next
End Function

Function LnoItrOfPubZ(M As CodeModule)
Asg Itr(LnoAyOfPubZ(M)), LnoItrOfPubZ
End Function


Private Function MthLinzEnsMdy$(OldMthLin, ShtMdy$)
Const CSub$ = CMod & "MthLinzEnsMdy"
Dim L$: L = RmvMdy(OldMthLin)
    Select Case ShtMdy
    Case "Pub", "": MthLinzEnsMdy = L
    Case "Prv":     MthLinzEnsMdy = "Private " & L
    Case "Frd":     MthLinzEnsMdy = "Friend " & L
    Case Else
        Thw CSub, "Given parameter [ShtMdy] must be ['' Pub Prv Frd]", "ShtMdy", ShtMdy
    End Select
End Function


