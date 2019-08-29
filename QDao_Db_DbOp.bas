Attribute VB_Name = "QDao_Db_DbOp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_Run."
Private Const Asm$ = "QDao"
Sub RunSqy(D As Database, Sqy$())
Dim Q$, I
For Each I In Sqy
    Q = I
    Rq D, Q
Next
End Sub

'
