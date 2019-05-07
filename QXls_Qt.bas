Attribute VB_Name = "QXls_Qt"
Option Explicit
Private Const CMod$ = "MXls_Qt."
Private Const Asm$ = "QXls"
Function FbtStrQt$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
FbtStrQt = FmtQQ("[?].[?]", DtaSrczScvl(CnStr), Tbl)
End Function
