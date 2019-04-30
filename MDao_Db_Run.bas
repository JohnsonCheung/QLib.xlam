Attribute VB_Name = "MDao_Db_Run"
Option Explicit
Sub RunSqy(A As Database, Sqy$())
Dim Q$, I
For Each I In Sqy
    Q = I
    RunQ A, Q
Next
End Sub
