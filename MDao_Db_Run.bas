Attribute VB_Name = "MDao_Db_Run"
Option Explicit
Sub RunSqy(A As Database, Sqy$())
Dim Q
For Each Q In Sqy
    RunQ A, Q
Next
End Sub
