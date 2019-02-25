Attribute VB_Name = "MDao_Def_Lin"
Option Explicit

Function FdStrAyFds(A As DAO.Fields) As String()
Dim F As DAO.Field
For Each F In A
    PushI FdStrAyFds, FdStr(F)
Next
End Function

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
