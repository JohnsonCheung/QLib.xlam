Attribute VB_Name = "QIde_ConstMth__Fun"
Option Explicit
Private Const CMod$ = "MIde_ConstMth__Fun."
Private Const Asm$ = "QIde"
Function FtzCnstQNm$(CnstQNm$)
Dim MdNm$, Nm$
FtzCnstQNm = ConstPrpPth(MdNm) & Nm & ".txt"
End Function
Private Function ConstPrpPth$(MdNm$)
ConstPrpPth = AddFdrEns(TmpHom, "ConstPrp", MdNm)
End Function

Function IsMthLinzConstStr(Lin$) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If BetBkt(Lin) <> "" Then Exit Function
If TakTyChr(Lin) = "$" Then Exit Function
IsMthLinzConstStr = True
End Function
Function IsMthLinzConstLy(Lin$) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If BetBkt(Lin) <> "" Then Exit Function
If MthRetTy(Lin) <> "String()" Then Exit Function
IsMthLinzConstLy = True
End Function

