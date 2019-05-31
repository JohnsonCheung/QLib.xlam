Attribute VB_Name = "QIde_Mth_Chr"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Chr."
Private Const Asm$ = "QIde"
Public Const TyChrLis$ = "!@#$%^&"

Function IsTyChr(A) As Boolean
If Len(A) <> 1 Then Exit Function
IsTyChr = HasSubStr(TyChrLis, A)
End Function
Function TyChrzTyNm$(TyNm)
Select Case TyNm
Case "String":   TyChrzTyNm = "$"
Case "Integer":  TyChrzTyNm = "%"
Case "Long":     TyChrzTyNm = "&"
Case "Double":   TyChrzTyNm = "#"
Case "Single":   TyChrzTyNm = "!"
Case "Currency": TyChrzTyNm = "@"
End Select
End Function

Function TyNmzTyChr$(TyChr$)
Dim O$
Select Case TyChr
Case "": Thw CSub, "TyChr cannot be blank"
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Thw CSub, "Invalid TyChr", "TyChr VdtTyChrLis", TyChr, TyChrLis
End Select
TyNmzTyChr = O
End Function

Function RmvTyChr$(S)
RmvTyChr = RmvChrzSfx(S, TyChrLis)
End Function

Function ShfTyChr$(OLin$)
ShfTyChr = ShfChr(OLin, TyChrLis)
End Function

Function TyChr$(Lin)
If IsMthLin(Lin) Then TyChr = TakTyChr(RmvMthn3(Lin))
End Function

Function TakTyChr$(S)
TakTyChr = TakChr(S, TyChrLis)
End Function
