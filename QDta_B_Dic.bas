Attribute VB_Name = "QDta_B_Dic"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Dic."
Private Const Asm$ = "QDta"

Function DoDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val") As Drs
DoDic = Drs(FoDic(InclDicValOptTy), DyoDic(A, InclDicValOptTy))
End Function

Function DtzDic(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValOptTy As Boolean) As DT
Dim Dy()
Dy = DyoDic(A, InclDicValOptTy)
Dim F$
    If InclDicValOptTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
DtzDic = DtzFF(DtNm, F, Dy)
End Function

Function FoDic(Optional InclValTy As Boolean) As String()
FoDic = SyzSS("Key Val" & IIf(InclValTy, " Type", ""))
End Function
