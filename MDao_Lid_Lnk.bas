Attribute VB_Name = "MDao_Lid_Lnk"
Option Explicit
Function ErzLnkTblzLtPm(A As Database, A() As LtPm) As String()
Dim J%
For J = 0 To UB(A)
    With A(J)
        PushIAy ErzLnkTblzLtPm, ErzLnkTblzTSrcCn(Db, .T, .S, .Cn)
    End With
Next
End Function
Sub LnkTblzLtPm(A As Database, A() As LtPm)
Dim J%
For J = 0 To UB(A)
    With A(J)
        LnkTbl Db, .T, .S, .Cn
    End With
Next
End Sub
Function TdzTSCn(T, Src, Cn) As DAO.TableDef
Set TdzTSCn = New DAO.TableDef
With TdzTSCn
    .Connect = Cn
    .Name = T
    .SourceTableName = Src
End With
End Function

Function LnkTny(A As Database) As String()
Dim T As TableDef
For Each T In A.TableDefs
    If T.Connect <> "" Then
        PushI LnkTny, T.Name
    End If
Next
End Function



