Attribute VB_Name = "QDao_Db_Run"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_Run."
Private Const Asm$ = "QDao"
Function RunSqy(A As Database, Sqy$()) As Unt
Dim Q$, I
For Each I In Sqy
    Q = I
    RunQ A, Q
Next
End Function
