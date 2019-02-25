Attribute VB_Name = "MDao_Db_Run"
Option Explicit
Sub RunSqyz(A As Database, SqlAy$())
Dim Q
For Each Q In SqlAy
    RunQz A, Q
Next
End Sub


Sub RunSqy(Sqy$())
RunSqyz CDb, Sqy
End Sub

