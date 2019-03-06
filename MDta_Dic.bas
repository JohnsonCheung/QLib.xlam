Attribute VB_Name = "MDta_Dic"
Option Explicit
Function DrszDic(A As Dictionary, Optional InclDicValOptTy As Boolean, Optional Tit$ = "Key Val") As Drs
Dim Fny$()
Fny = SySsl(Tit): If InclDicValOptTy Then Push Fny, "Val-TypeName"
Set DrszDic = Drs(Fny, DryzDic(A, InclDicValOptTy))
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
Set DtzDic = Dt(DtNm, F, Dry)
End Function

Function FnyzDic(Optional InclValTy As Boolean) As String()
FnyzDic = SySsl("Key Val" & IIf(InclValTy, " Type", ""))
End Function
