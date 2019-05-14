Attribute VB_Name = "QDao_Lid_Lnk"
Option Explicit
Private Const CMod$ = "MDao_Lid_Lnk."
Private Const Asm$ = "QDao"
Function ErzLnkTblPms(A As Database, B As LnkTblPms) As String()
Dim J%, Ay() As LnkTblPm
Ay = B.Ay
For J = 0 To B.N - 1
    With Ay(J)
        PushIAy ErzLnkTblPms, ErzLnkTblzTSrcCn(A, .T, .S, .Cn)
    End With
Next
End Function
Function TdzTSCn(T, Src$, Cn$) As Dao.TableDef
Set TdzTSCn = New Dao.TableDef
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



