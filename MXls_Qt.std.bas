Attribute VB_Name = "MXls_Qt"
Option Explicit
Function FbtStrQt$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
FbtStrQt = FmtQQ("[?].[?]", DtaSrczTdCn(CnStr), Tbl)
End Function
