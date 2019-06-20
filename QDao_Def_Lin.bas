Attribute VB_Name = "QDao_Def_Lin"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Def_Lin."
Private Const Asm$ = "QDao"

Function IdxStr$(A As DAO.Index)
Dim X$, F$
With A
IdxStr = FmtQQ("Idx;?;?;?", .Name, X, F)
End With
End Function

Function IdxStrAyIdxs(A As DAO.Indexes) As String()
Dim I As DAO.Index
For Each I In A
    PushI IdxStrAyIdxs, IdxStr(I)
Next
End Function
