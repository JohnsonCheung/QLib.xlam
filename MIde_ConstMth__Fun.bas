Attribute VB_Name = "MIde_ConstMth__Fun"
Option Explicit
Function FtzConstQNm$(ConstQNm$)
Dim MdNm$, Nm$
FtzConstQNm = ConstMthPth(MdNm) & Nm & ".txt"
End Function
Private Function ConstMthPth$(MdNm$)
ConstMthPth = AddFdrEns(TmpHom, "ConstMth", MdNm)
End Function

Function IsMthLinzConstStr(Lin) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If StrBetBkt(Lin) <> "" Then Exit Function
If TakMthChr(Lin) = "$" Then Exit Function
IsMthLinzConstStr = True
End Function
Function IsMthLinzConstLy(Lin) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If StrBetBkt(Lin) <> "" Then Exit Function
If TakMthRetTy(Lin) <> "As String()" Then Exit Function
IsMthLinzConstLy = True
End Function

