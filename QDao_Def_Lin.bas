Attribute VB_Name = "QDao_Def_Lin"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Def_Lin."
Private Const Asm$ = "QDao"

Function IdxStr$(A As Dao.Index)
Dim X$, F$
With A
IdxStr = FmtQQ("Idx;?;?;?", .Name, X, F)
End With
End Function

Function IdxStrAyIdxs(A As Dao.Indexes) As String()
Dim I As Dao.Index
For Each I In A
    PushI IdxStrAyIdxs, IdxStr(I)
Next
End Function
