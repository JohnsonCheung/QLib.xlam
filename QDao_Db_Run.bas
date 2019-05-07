Attribute VB_Name = "QDao_Db_Run"
Option Explicit
Private Const CMod$ = "MDao_Db_Run."
Private Const Asm$ = "QDao"
Sub RunSqy(A As Database, Sqy$())
Dim Q$, I
For Each I In Sqy
    Q = I
    RunQ A, Q
Next
End Sub
