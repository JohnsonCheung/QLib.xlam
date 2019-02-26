Attribute VB_Name = "MDao_Def_Lin"
Option Explicit

Function FdStrAyFds(A As Dao.Fields) As String()
Dim F As Dao.Field
For Each F In A
    PushI FdStrAyFds, FdStr(F)
Next
End Function

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
