Attribute VB_Name = "QIde_Cd_Cd"
Function TyCd$(Tyn, Optional IsPrv As Boolean)
Dim C1$: C1 = TpOfTys(IsPrv)
Dim C2$: C2 = TpOfPushTy(IsPrv)
Dim C3$: C3 = TpOfPushTys(IsPrv)
Dim C4$: C4 = TpOfAddTy(IsPrv)
Dim C5$: C5 = TpOfSngTy(IsPrv)
TyCd = RplQ(JnCrLf(Sy(C1, C2, C3, C4, C5)), Tyn)
End Function

Private Function TpOfPushTy$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Push?(O As ?s, M As ?)"
X "ReDim Preserve O.Ay(O.N)"
X "O.Ay(O.N) = M"
X "O.N = O.N + 1"
X "End Sub"
TpOfPushTy = JnCrLf(XX)
Erase XX
End Function

Private Function TpOfPushTys$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Push?s(O As ?s, M As ?s)"
X "Dim J&"
X "For J=0 To ?.N - 1"
X "    Push? O, A.Ay(J)"
X "Next"
X "End Sub"
TpOfPushTys = JnCrLf(XX)
Erase XX
End Function

Private Function Prv$(IsPrv As Boolean)
If IsPrv Then Prv = "Private "
End Function

Private Function TpOfTys$(Optional IsPrv As Boolean)
TpOfTys = vbCrLf & Prv(IsPrv) & "Type ?s: N As Long: Ay() As ?: End Type"
End Function

Private Function TpOfAddTy$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Add?(A As ?, B As ?) As ?s"
X "Push? Add?, A"
X "Push? Add?, B"
X "End Sub"
TpOfAddTy = JnCrLf(XX)
Erase XX
End Function

Function TpOfSngTy$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Sng?(A As ?) As ?s"
X "Push? Sng?, A"
X "End Sub"
TpOfSngTy = JnCrLf(XX)
Erase XX
End Function

