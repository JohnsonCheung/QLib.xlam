Attribute VB_Name = "QDta_Dic"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Dic."
Private Const Asm$ = "QDta"
Function DrszDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val") As Drs
Dim Fny$()
Fny = SyzSS(Tit): If InclDicValOptTy Then Push Fny, "Val-TypeName"
DrszDic = Drs(Fny, DryzDic(A, InclDicValOptTy))
End Function

Function DtzDic(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValOptTy As Boolean) As Dt
Dim Dry()
Dry = DryzDic(A, InclDicValOptTy)
Dim F$
    If InclDicValOptTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
DtzDic = DtzFF(DtNm, F, Dry)
End Function

Function FnyzDic(Optional InclValTy As Boolean) As String()
FnyzDic = SyzSS("Key Val" & IIf(InclValTy, " Type", ""))
End Function
