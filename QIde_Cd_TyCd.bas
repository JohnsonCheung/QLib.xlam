Attribute VB_Name = "QIde_Cd_TyCd"
Option Explicit
Option Compare Text
Function TyCd$(Tyn, Optional IsPrv As Boolean)
Dim C1$: C1 = Tp_Tys(IsPrv)
Dim C2$: C2 = Tp_Push(IsPrv)
Dim C3$: C3 = Tp_Pushs(IsPrv)
Dim C4$: C4 = Tp_Add(IsPrv)
Dim C5$: C5 = Tp_Sng(IsPrv)
TyCd = SzQBy(JnCrLf(Sy(C1, C2, C3, C4, C5)), Tyn)
End Function

Function TyCd_Tys$(Tyn, Optional IsPrv As Boolean)
TyCd_Tys = SzQBy(Tp_Tys(IsPrv), Tyn)
End Function

Private Function Tp_Push$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Push?(O As ?s, M As ?)"
X "ReDim Preserve O.Ay(O.N)"
X "O.Ay(O.N) = M"
X "O.N = O.N + 1"
X "End Sub"
Tp_Push = JnCrLf(XX)
Erase XX
End Function

Private Function Tp_Pushs$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Push?s(O As ?s, M As ?s)"
X "Dim J&"
X "For J=0 To ?.N - 1"
X "    Push? O, A.Ay(J)"
X "Next"
X "End Sub"
Tp_Pushs = JnCrLf(XX)
Erase XX
End Function

Private Function Prv$(IsPrv As Boolean)
If IsPrv Then Prv = "Private "
End Function

Private Function Tp_Tys$(Optional IsPrv As Boolean)
Tp_Tys = vbCrLf & Prv(IsPrv) & "Type ?s: N As Long: Ay() As ?: End Type"
End Function

Private Function Tp_Add$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Add?(A As ?, B As ?) As ?s"
X "Push? Add?, A"
X "Push? Add?, B"
X "End Sub"
Tp_Add = JnCrLf(XX)
Erase XX
End Function

Function Tp_Sng$(Optional IsPrv As Boolean)
Erase XX
X ""
X Prv(IsPrv) & "Sub Sng?(A As ?) As ?s"
X "Push? Sng?, A"
X "End Sub"
Tp_Sng = JnCrLf(XX)
Erase XX
End Function

