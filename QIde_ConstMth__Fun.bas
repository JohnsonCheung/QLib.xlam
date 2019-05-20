Attribute VB_Name = "QIde_ConstMth__Fun"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_ConstMth__Fun."
Private Const Asm$ = "QIde"
Function FtzCnstQNm$(CnstQNm$)
Dim Mdn, Nm$
FtzCnstQNm = ConstPrpPth(Mdn) & Nm & ".txt"
End Function
Private Function ConstPrpPth$(Mdn)
'ConstPrpPth = AddFdrEns(TmpHom, "ConstPrp", Mdn)
End Function

Function IsMthLinzConstStr(Lin) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If BetBkt(Lin) <> "" Then Exit Function
If TakTyChr(Lin) = "$" Then Exit Function
IsMthLinzConstStr = True
End Function
Function IsMthLinzConstLy(Lin) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If BetBkt(Lin) <> "" Then Exit Function
If MthRetTy(Lin) <> "String()" Then Exit Function
IsMthLinzConstLy = True
End Function

